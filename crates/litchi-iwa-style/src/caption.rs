//! The inline title and caption graph of a drawable (records 0371, 0372,
//! 0425), shared by the Keynote, Pages and Numbers owners.
//!
//! A drawable's title and caption are references in its
//! `TSD.DrawableArchive` (fields 10 and 11). A present label is a
//! `TSA.CaptionInfoArchive` with a private text storage, a shape style and
//! a placement in the drawable's component; an absent label is an empty
//! `TSD.StandinCaptionArchive`. Setting a present label rewrites its
//! storage text; setting an absent one authors the four-object inline graph
//! (the shared `keynote_chart_caption_graph_codec`, the payloads the host
//! writes) and points the drawable at it; clearing a present label points
//! the drawable at a fresh stand-in and keeps the old graph, as the
//! application does for undo.
//!
//! Two layers: payload- and object-level rules ([`Label`], [`GraphIds`],
//! [`graph_objects`], [`standin`], [`info_graph`], [`repoint`],
//! [`replace_reference`], [`drawable_width`], [`theme_paragraph_style`])
//! for owners that publish through their own cores (Keynote, Pages); and
//! [`read()`] and [`write()`] over a [`Workspace`] for owners that publish a
//! workspace's changes (Numbers).

use litchi_iwa_common::WireLimits;
use litchi_iwa_common::wire::{
    WireView, patch_length_delimited_field, transform_length_delimited_field,
};
use litchi_iwa_core::{ArchiveObject, RawMessage};
use litchi_iwa_protos::{
    keynote_chart_caption_graph_codec as graph_codec, pages_movie_caption_codec,
};

use crate::graph::Workspace;

/// `TSA.CaptionInfoArchive`.
pub const CAPTION_INFO_MESSAGE_TYPE: u32 = 633;
/// `TSA.CaptionPlacementArchive`.
pub const CAPTION_PLACEMENT_MESSAGE_TYPE: u32 = 634;
/// `TSWP.ShapeStyleArchive`.
pub const SHAPE_STYLE_MESSAGE_TYPE: u32 = 2_025;
/// `TSWP.StorageArchive`.
pub const STORAGE_MESSAGE_TYPE: u32 = 2_001;
/// `TSD.StandinCaptionArchive`.
pub const STANDIN_MESSAGE_TYPE: u32 = 3_097;
/// The message version the host writes on a new label graph's objects.
pub const GRAPH_VERSION: [u32; 3] = [1, 0, 5];
/// The message version the host writes on a new stand-in.
pub const STANDIN_VERSION: [u32; 3] = [10, 1, 0];

const DRAWABLE_GEOMETRY_FIELD: u32 = 1;
const GEOMETRY_SIZE_FIELD: u32 = 2;
const SIZE_WIDTH_FIELD: u32 = 1;
const THEME_PRESETS_FIELD: u32 = 210;
const PRESETS_CAPTION_STYLES_FIELD: u32 = 1;

/// A malformed drawable payload or label reference.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Malformed;

/// Which of a drawable's two labels.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Label {
    /// The label above the drawable (`TSD.DrawableArchive.title`).
    Title,
    /// The label below the drawable (`TSD.DrawableArchive.caption`).
    Caption,
}

impl Label {
    /// The `TSD.DrawableArchive` field of the label's reference.
    #[must_use]
    pub const fn field(self) -> u32 {
        match self {
            Self::Title => 10,
            Self::Caption => 11,
        }
    }

    /// `TSA.CaptionInfoArchive.child_info_kind` of the label.
    #[must_use]
    pub const fn native_kind(self) -> i32 {
        match self {
            Self::Caption => 1,
            Self::Title => 2,
        }
    }

    /// The other label of the drawable.
    #[must_use]
    pub const fn other(self) -> Self {
        match self {
            Self::Title => Self::Caption,
            Self::Caption => Self::Title,
        }
    }

    const fn graph_kind(self) -> graph_codec::CaptionGraphKind {
        match self {
            Self::Caption => graph_codec::CaptionGraphKind::Caption,
            Self::Title => graph_codec::CaptionGraphKind::Title,
        }
    }
}

/// Where a drawable message stores its `TSD.DrawableArchive`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Nesting {
    /// Field 1 (images, movies, charts).
    Drawable,
    /// Field 1 of field 1 (`TSWP.ShapeInfoArchive.super.super`: shapes and
    /// text boxes).
    Shape,
}

/// The four objects of a new label graph.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct GraphIds {
    /// The label's shape style.
    pub style: u64,
    /// The `TSA.CaptionInfoArchive` the drawable references.
    pub info: u64,
    /// The text storage.
    pub storage: u64,
    /// The placement.
    pub placement: u64,
}

impl GraphIds {
    /// Four consecutive identifiers from `first` (the host's order: style,
    /// info, storage, placement); `None` on overflow.
    #[must_use]
    pub fn allocate(first: u64) -> Option<Self> {
        Some(Self {
            style: first,
            info: first.checked_add(1)?,
            storage: first.checked_add(2)?,
            placement: first.checked_add(3)?,
        })
    }

    /// The identifiers in allocation order.
    #[must_use]
    pub const fn all(self) -> [u64; 4] {
        [self.style, self.info, self.storage, self.placement]
    }

    /// The last allocated identifier (the new watermark).
    #[must_use]
    pub const fn last(self) -> u64 {
        self.placement
    }
}

/// The theme objects and language a new label graph references.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Theme<'a> {
    /// The document stylesheet.
    pub stylesheet: u64,
    /// The theme's first caption paragraph-style preset.
    pub paragraph_style: u64,
    /// The document language.
    pub language: Option<&'a str>,
}

/// A failure building label objects.
#[derive(Debug)]
pub enum BuildError {
    /// The graph codec refused or exhausted a limit.
    Encode(graph_codec::EncodeError),
    /// An object could not be built.
    Archive(litchi_iwa_core::Error),
}

fn versioned(
    identifier: u64,
    message_type: u32,
    data: Vec<u8>,
    version: [u32; 3],
    references: Vec<u64>,
    limits: litchi_iwa_core::Limits,
) -> Result<ArchiveObject, litchi_iwa_core::Error> {
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: message_type,
            data,
        }],
        limits,
    )?;
    let info = &mut object.archive_info.message_infos[0];
    info.versions = version.to_vec().into();
    info.object_references = references;
    Ok(object)
}

/// The four objects of a new inline label graph of `drawable` (shape
/// style, caption info, storage, placement; the host's
/// `caption_objects`) under the archive `limits`, and the codec's report.
///
/// # Errors
///
/// The codec's refusals and limits, or an object that cannot be built.
#[allow(clippy::too_many_arguments)]
pub fn graph_objects(
    ids: GraphIds,
    drawable: u64,
    width: f32,
    text: &str,
    theme: Theme<'_>,
    label: Label,
    options: graph_codec::EncodeOptions,
    limits: litchi_iwa_core::Limits,
) -> Result<(Vec<ArchiveObject>, graph_codec::EncodeReport), BuildError> {
    let output = graph_codec::encode_caption_graph_with_kind_with_report(
        graph_codec::CaptionGraphWrite {
            drawable_identifier: drawable,
            style_identifier: ids.style,
            info_identifier: ids.info,
            storage_identifier: ids.storage,
            placement_identifier: ids.placement,
            stylesheet_identifier: theme.stylesheet,
            paragraph_style_identifier: theme.paragraph_style,
            drawable_width: width,
            text,
            language: theme.language,
        },
        label.graph_kind(),
        options,
    )
    .map_err(BuildError::Encode)?;
    let report = output.report();
    let [style, info, storage, placement] = output.into_payloads().into_parts();
    let objects = vec![
        versioned(
            ids.style,
            SHAPE_STYLE_MESSAGE_TYPE,
            style,
            GRAPH_VERSION,
            vec![theme.paragraph_style],
            limits,
        ),
        versioned(
            ids.info,
            CAPTION_INFO_MESSAGE_TYPE,
            info,
            GRAPH_VERSION,
            vec![ids.style, ids.storage, ids.placement],
            limits,
        ),
        versioned(
            ids.storage,
            STORAGE_MESSAGE_TYPE,
            storage,
            GRAPH_VERSION,
            vec![theme.paragraph_style],
            limits,
        ),
        versioned(
            ids.placement,
            CAPTION_PLACEMENT_MESSAGE_TYPE,
            placement,
            GRAPH_VERSION,
            Vec::new(),
            limits,
        ),
    ]
    .into_iter()
    .collect::<Result<Vec<_>, _>>()
    .map_err(BuildError::Archive)?;
    Ok((objects, report))
}

/// A canonical empty stand-in object (the host's `standin_caption_object`)
/// under the archive `limits`.
///
/// # Errors
///
/// An object that cannot be built.
pub fn standin(
    identifier: u64,
    limits: litchi_iwa_core::Limits,
) -> Result<ArchiveObject, litchi_iwa_core::Error> {
    versioned(
        identifier,
        STANDIN_MESSAGE_TYPE,
        graph_codec::canonical_standin_payload().to_vec(),
        STANDIN_VERSION,
        Vec::new(),
        limits,
    )
}

/// The host's `replace_object_reference` over a message's reference list:
/// `old` becomes `new` (a then-duplicate `new` is dropped), or `new` is
/// appended when `old` is not listed.
pub fn replace_reference(references: &mut Vec<u64>, old: u64, new: u64) {
    if let Some(index) = references.iter().position(|identifier| *identifier == old) {
        references[index] = new;
        let mut seen = false;
        references.retain(|identifier| {
            if *identifier != new {
                return true;
            }
            let keep = !seen;
            seen = true;
            keep
        });
    } else if !references.contains(&new) {
        references.push(new);
    }
}

/// The private objects of a present label (`TSA.CaptionInfoArchive`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct InfoGraph {
    /// The shape style.
    pub style: u64,
    /// The text storage.
    pub storage: u64,
    /// The placement.
    pub placement: u64,
}

/// The objects a caption info payload names, when it is the `label` of
/// `drawable`: a text-box storage owned (and deprecated-referenced) by the
/// info, a shape style and a placement, all distinct from `info` and each
/// other. `None` for any other shape.
#[must_use]
pub fn info_graph(payload: &[u8], info: u64, drawable: u64, label: Label) -> Option<InfoGraph> {
    let snapshot = pages_movie_caption_codec::decode_caption_info(
        payload,
        pages_movie_caption_codec::DecodeOptions::for_source(payload),
    )
    .ok()?;
    let storage = snapshot.owned_storage_identifier().filter(|id| *id != 0)?;
    if snapshot.child_info_kind() != Some(label.native_kind())
        || snapshot.parent_identifier() != drawable
        || snapshot.deprecated_storage_identifier() != Some(storage)
        || snapshot.is_text_box() != Some(true)
    {
        return None;
    }
    let style = snapshot.style_identifier().filter(|id| *id != 0)?;
    let placement = snapshot.placement_identifier().filter(|id| *id != 0)?;
    let ids = [info, style, storage, placement];
    for (index, id) in ids.iter().enumerate() {
        if ids[index + 1..].contains(id) {
            return None;
        }
    }
    Some(InfoGraph {
        style,
        storage,
        placement,
    })
}

/// A `TSP.Reference` payload naming `identifier`.
#[must_use]
pub fn reference_payload(identifier: u64) -> Vec<u8> {
    let mut payload = vec![0x08];
    litchi_iwa_common::varint::encode_varint_into(&mut payload, identifier);
    payload
}

/// The drawable message `payload` with its `label` reference pointing at
/// `replacement` (every other byte kept).
///
/// # Errors
///
/// An invalid payload.
pub fn repoint(
    payload: &[u8],
    nesting: Nesting,
    label: Label,
    replacement: u64,
) -> litchi_iwa_common::Result<Vec<u8>> {
    let reference = reference_payload(replacement);
    let patch = |drawable: &[u8]| {
        patch_length_delimited_field(drawable, label.field(), true, Some(&reference))
    };
    match nesting {
        Nesting::Shape => transform_length_delimited_field(payload, 1, |shape| {
            transform_length_delimited_field(shape, 1, patch)
        }),
        Nesting::Drawable => transform_length_delimited_field(payload, 1, patch),
    }
}

/// The single length-delimited field `number` of `view`; `Err` for a
/// repeated field or another wire type.
fn unique<'a>(view: &WireView<'a>, number: u32) -> Result<Option<&'a [u8]>, Malformed> {
    let mut found = None;
    for field in view.fields().filter(|field| field.number() == number) {
        if found.is_some() || field.wire_type() != 2 {
            return Err(Malformed);
        }
        found = Some(field.payload());
    }
    Ok(found)
}

/// The `TSD.DrawableArchive` of a drawable message payload.
fn drawable_archive(payload: &[u8], nesting: Nesting) -> Result<&[u8], Malformed> {
    let limits = WireLimits::DEFAULT;
    let outer = WireView::parse(payload, limits).map_err(|_| Malformed)?;
    let mut drawable = unique(&outer, 1)?.ok_or(Malformed)?;
    if nesting == Nesting::Shape {
        let shape = WireView::parse(drawable, limits).map_err(|_| Malformed)?;
        drawable = unique(&shape, 1)?.ok_or(Malformed)?;
    }
    Ok(drawable)
}

/// The identifier `label` references in a drawable message payload;
/// `Ok(None)` when the drawable stores no such reference.
///
/// # Errors
///
/// An invalid payload or reference.
pub fn label_reference(
    payload: &[u8],
    nesting: Nesting,
    label: Label,
) -> Result<Option<u64>, Malformed> {
    let drawable = drawable_archive(payload, nesting)?;
    let view = WireView::parse(drawable, WireLimits::DEFAULT).map_err(|_| Malformed)?;
    let Some(reference) = unique(&view, label.field())? else {
        return Ok(None);
    };
    let reference = WireView::parse(reference, WireLimits::DEFAULT).map_err(|_| Malformed)?;
    let mut identifier = None;
    for field in reference.fields().filter(|field| field.number() == 1) {
        if identifier.is_some() || field.wire_type() != 0 {
            return Err(Malformed);
        }
        let (value, _) =
            litchi_iwa_common::decode_varint_from_bytes(field.payload()).map_err(|_| Malformed)?;
        identifier = Some(value);
    }
    match identifier {
        Some(0) | None => Err(Malformed),
        Some(value) => Ok(Some(value)),
    }
}

/// The displayed width (`geometry.size.width`) of a drawable message
/// payload; `Ok(None)` when it stores none.
///
/// # Errors
///
/// An invalid payload, or a repeated or mistyped width.
pub fn drawable_width(payload: &[u8], nesting: Nesting) -> Result<Option<f32>, Malformed> {
    archive_width(drawable_archive(payload, nesting)?)
}

/// The displayed width of a `TSD.DrawableArchive` payload; `Ok(None)` when
/// it stores none.
///
/// # Errors
///
/// An invalid payload, or a repeated or mistyped width.
pub fn archive_width(drawable: &[u8]) -> Result<Option<f32>, Malformed> {
    let limits = WireLimits::DEFAULT;
    let view = WireView::parse(drawable, limits).map_err(|_| Malformed)?;
    let Some(geometry) = unique(&view, DRAWABLE_GEOMETRY_FIELD)? else {
        return Ok(None);
    };
    let geometry = WireView::parse(geometry, limits).map_err(|_| Malformed)?;
    let Some(size) = unique(&geometry, GEOMETRY_SIZE_FIELD)? else {
        return Ok(None);
    };
    let size = WireView::parse(size, limits).map_err(|_| Malformed)?;
    let mut width = None;
    for field in size
        .fields()
        .filter(|field| field.number() == SIZE_WIDTH_FIELD)
    {
        if width.is_some() || field.wire_type() != 5 {
            return Err(Malformed);
        }
        let bytes: [u8; 4] = field.payload().try_into().map_err(|_| Malformed)?;
        width = Some(f32::from_le_bytes(bytes));
    }
    Ok(width)
}

/// The first caption paragraph-style preset of a `TSS.ThemeArchive`
/// payload (`application_presets.caption_styles[0]`, the host's
/// `image_caption_theme`).
#[must_use]
pub fn theme_paragraph_style(theme: &[u8]) -> Option<u64> {
    let limits = WireLimits::DEFAULT;
    let view = WireView::parse(theme, limits).ok()?;
    let presets = unique(&view, THEME_PRESETS_FIELD).ok()??;
    let presets = WireView::parse(presets, limits).ok()?;
    let first = presets
        .fields()
        .find(|field| field.number() == PRESETS_CAPTION_STYLES_FIELD)?;
    if first.wire_type() != 2 {
        return None;
    }
    let reference = WireView::parse(first.payload(), limits).ok()?;
    let field = reference.fields().find(|field| field.number() == 1)?;
    let (value, _) = litchi_iwa_common::decode_varint_from_bytes(field.payload()).ok()?;
    (value != 0).then_some(value)
}

/// One label of a drawable in a [`Workspace`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Target {
    /// The drawable object.
    pub drawable: u64,
    /// Where its message stores the `TSD.DrawableArchive`.
    pub nesting: Nesting,
    /// The label.
    pub label: Label,
}

/// A refusal or failure of the workspace layer.
#[derive(Debug)]
pub enum Error {
    /// The drawable or label graph is malformed (or another graph failure).
    Graph(crate::graph::Error),
    /// Authoring a new label needs a positive displayed width.
    NoWidth,
    /// Replacing the present label's text would drop an object its text
    /// references.
    DependentContent,
    /// The graph codec refused or exhausted a limit.
    Encode(graph_codec::EncodeError),
    /// The text storage rewrite refused or exhausted a limit.
    Storage,
}

impl From<crate::graph::Error> for Error {
    fn from(error: crate::graph::Error) -> Self {
        Self::Graph(error)
    }
}

fn invalid(reason: &'static str) -> Error {
    Error::Graph(crate::graph::Error::Invalid(reason))
}

/// The resolved state of one label.
#[derive(Debug, Clone, PartialEq)]
pub struct Slot {
    /// The object the drawable references (a caption info or stand-in).
    pub reference: u64,
    /// The present label's graph; `None` for a stand-in.
    pub graph: Option<InfoGraph>,
    /// The present label's text.
    pub text: Option<String>,
    /// The drawable's displayed width.
    pub width: Option<f32>,
}

fn drawable_message<'w>(
    workspace: &'w Workspace<'_>,
    drawable: u64,
) -> Result<(usize, &'w [u8]), Error> {
    let object = workspace.object(drawable)?;
    let [message] = object.messages.as_slice() else {
        return Err(invalid("a drawable must have exactly one message"));
    };
    Ok((0, message.data.as_slice()))
}

fn storage_text(payload: &[u8]) -> Result<String, Error> {
    let storage = litchi_iwa_text_wire::from_bytes(
        payload,
        litchi_iwa_text_wire::Limits::new(
            litchi_iwa_text_wire::DEFAULT_MAX_MESSAGE_BYTES,
            litchi_iwa_text_wire::DEFAULT_MAX_FIELDS,
            litchi_iwa_text_wire::DEFAULT_MAX_WIRE_FRAGMENTS,
            litchi_iwa_text_wire::DEFAULT_MAX_TEXT_BYTES,
        )
        .map_err(|_| invalid("text storage limits"))?,
    )
    .map_err(|_| invalid("an invalid label storage"))?;
    Ok(storage.into_text())
}

/// Resolve `target` (the host's `caption_slot_from_reference`): the
/// referenced object must share the drawable's component and be either one
/// stand-in or a caption info whose storage and placement share it too.
///
/// # Errors
///
/// A malformed drawable or label graph.
pub fn resolve(workspace: &Workspace<'_>, target: Target) -> Result<Slot, Error> {
    let (_, payload) = drawable_message(workspace, target.drawable)?;
    let reference = label_reference(payload, target.nesting, target.label)
        .map_err(|Malformed| invalid("an invalid drawable label reference"))?
        .ok_or(invalid("a drawable without a label reference"))?;
    let width = drawable_width(payload, target.nesting)
        .map_err(|Malformed| invalid("an invalid drawable width"))?;
    let home = workspace.locate(target.drawable)?;
    if workspace.locate(reference)? != home {
        return Err(invalid("a label outside its drawable's component"));
    }
    let object = workspace.object(reference)?;
    let infos = object
        .messages
        .iter()
        .filter(|message| message.type_ == CAPTION_INFO_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    let [info] = infos.as_slice() else {
        if !infos.is_empty() {
            return Err(invalid("a label with several caption infos"));
        }
        let [message] = object.messages.as_slice() else {
            return Err(invalid("a stand-in label with several messages"));
        };
        if message.type_ != STANDIN_MESSAGE_TYPE {
            return Err(invalid("a label that is neither a caption nor a stand-in"));
        }
        return Ok(Slot {
            reference,
            graph: None,
            text: None,
            width,
        });
    };
    let graph = info_graph(&info.data, reference, target.drawable, target.label)
        .ok_or(invalid("an invalid caption info"))?;
    for (identifier, message_type, private) in [
        (graph.style, SHAPE_STYLE_MESSAGE_TYPE, false),
        (graph.storage, STORAGE_MESSAGE_TYPE, true),
        (graph.placement, CAPTION_PLACEMENT_MESSAGE_TYPE, true),
    ] {
        let object = workspace.object(identifier)?;
        let [message] = object.messages.as_slice() else {
            return Err(invalid("a label object with several messages"));
        };
        if message.type_ != message_type {
            return Err(invalid("a label object of another type"));
        }
        if private && workspace.locate(identifier)? != home {
            return Err(invalid("a label object outside its drawable's component"));
        }
    }
    let storage = workspace.object(graph.storage)?;
    let text = storage_text(&storage.messages[0].data)?;
    Ok(Slot {
        reference,
        graph: Some(graph),
        text: Some(text),
        width,
    })
}

/// Read `target`'s text; `None` when the label is absent.
///
/// # Errors
///
/// A malformed drawable or label graph.
pub fn read(workspace: &Workspace<'_>, target: Target) -> Result<Option<String>, Error> {
    Ok(resolve(workspace, target)?.text)
}

/// Whether the component storing `object` also stores an object with a
/// message of `message_type` (Numbers stores a chart drawn in the document
/// component, beside the calculation engine, with the host's componentized
/// caption layout).
///
/// # Errors
///
/// A missing object.
pub fn component_holds(
    workspace: &Workspace<'_>,
    object: u64,
    message_type: u32,
) -> Result<bool, Error> {
    let archive = workspace.archive_at(workspace.locate(object)?)?;
    Ok(archive.objects.iter().any(|object| {
        object
            .messages
            .iter()
            .any(|message| message.type_ == message_type)
    }))
}

/// The objects a [`write()`] created.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Written {
    /// Whether anything changed.
    pub changed: bool,
    /// The new objects (a label graph or a stand-in).
    pub created: Vec<u64>,
}

fn point_drawable_at(
    workspace: &mut Workspace<'_>,
    target: Target,
    old: u64,
    new: u64,
) -> Result<(), Error> {
    let (message, payload) = drawable_message(workspace, target.drawable)?;
    let data = repoint(payload, target.nesting, target.label, new)
        .map_err(|error| Error::Graph(crate::graph::Error::Wire(error)))?;
    workspace.replace_payload(target.drawable, message, data)?;
    let info = workspace
        .object_mut(target.drawable)?
        .archive_info
        .message_infos
        .get_mut(message)
        .ok_or(invalid("a drawable without a message info"))?;
    replace_reference(&mut info.object_references, old, new);
    Ok(())
}

/// Store `text` as `target`'s label (`None`: clear it), as the host's
/// `set_*_caption`/`remove_*_caption` do: a present label's storage text
/// is replaced; an absent one gets a new inline graph (`theme` names its
/// stylesheet, paragraph style and language); clearing points the drawable
/// at a fresh stand-in and keeps the old graph. New objects take
/// identifiers from the package watermark, get UUIDs in the drawable's
/// component and raise the watermark.
///
/// # Errors
///
/// [`Error::NoWidth`], [`Error::DependentContent`], a malformed graph, or
/// a codec or storage failure.
pub fn write(
    workspace: &mut Workspace<'_>,
    target: Target,
    text: Option<&str>,
    theme: impl FnOnce(&Workspace<'_>) -> Result<(u64, u64, Option<String>), Error>,
) -> Result<Written, Error> {
    let slot = resolve(workspace, target)?;
    if slot.text.as_deref() == text {
        return Ok(Written::default());
    }
    let member = workspace
        .member(workspace.locate(target.drawable)?)
        .to_owned();
    let component = workspace
        .component_for_member(&member)?
        .ok_or(invalid("a drawable outside every component"))?;
    let location = workspace.locate(target.drawable)?;
    let created = match (text, slot.graph) {
        (Some(text), Some(graph)) => {
            let storage = workspace.object(graph.storage)?;
            let payload = storage.messages[0].data.clone();
            let length = slot
                .text
                .as_deref()
                .map_or(0, |text| text.encode_utf16().count());
            let prepared = litchi_iwa_text_wire::prepare_storage_text_rewrite_with_behavior(
                &payload,
                0..length,
                text,
                litchi_iwa_text_wire::RewriteBehavior::ReplaceSelection,
                litchi_iwa_text_wire::RewriteLimits::default(),
            )
            .map_err(|_| Error::Storage)?;
            let requirements = prepared.execution_requirements();
            let rewritten = prepared
                .execute(litchi_iwa_text_wire::StorageRewriteExecutionLimits {
                    max_output_bytes: requirements.output_bytes(),
                    max_retained_elements: requirements.retained_elements(),
                    max_retained_bytes: requirements.retained_bytes(),
                    max_peak_scratch_bytes: requirements.peak_scratch_bytes(),
                    max_allocations: requirements.allocations(),
                    max_work: requirements.work(),
                })
                .map_err(|_| Error::Storage)?;
            if !rewritten.removed_object_references().is_empty() {
                return Err(Error::DependentContent);
            }
            workspace.replace_payload(graph.storage, 0, rewritten.into_bytes())?;
            Vec::new()
        },
        (Some(text), None) => {
            let width = slot
                .width
                .filter(|width| width.is_finite() && *width > 0.0)
                .ok_or(Error::NoWidth)?;
            let (stylesheet, paragraph_style, language) = theme(workspace)?;
            let ids = GraphIds::allocate(workspace.next_object_identifier()?)
                .ok_or(Error::Graph(crate::graph::Error::IdentifierOverflow))?;
            let (objects, _) = graph_objects(
                ids,
                target.drawable,
                width,
                text,
                Theme {
                    stylesheet,
                    paragraph_style,
                    language: language.as_deref(),
                },
                target.label,
                graph_codec::EncodeOptions::for_text(text),
                litchi_iwa_core::Limits::default(),
            )
            .map_err(|error| match error {
                BuildError::Encode(error) => Error::Encode(error),
                BuildError::Archive(error) => Error::Graph(crate::graph::Error::Archive(error)),
            })?;
            for object in objects {
                workspace.insert(location, object)?;
            }
            point_drawable_at(workspace, target, slot.reference, ids.info)?;
            ids.all().to_vec()
        },
        (None, Some(_)) => {
            let identifier = workspace.next_object_identifier()?;
            workspace.insert(
                location,
                standin(identifier, litchi_iwa_core::Limits::default())
                    .map_err(crate::graph::Error::Archive)?,
            )?;
            point_drawable_at(workspace, target, slot.reference, identifier)?;
            vec![identifier]
        },
        (None, None) => return Ok(Written::default()),
    };
    if let Some(&last) = created.last() {
        workspace.add_object_uuids(component, &created)?;
        workspace.set_last_object_identifier(last)?;
    }
    if read(workspace, target)?.as_deref() != text {
        return Err(invalid("a label write does not read back"));
    }
    Ok(Written {
        changed: true,
        created,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn replace_reference_follows_the_host() {
        let mut references = vec![4, 9, 7];
        replace_reference(&mut references, 9, 7);
        assert_eq!(references, vec![4, 7]);
        replace_reference(&mut references, 3, 11);
        assert_eq!(references, vec![4, 7, 11]);
        replace_reference(&mut references, 3, 11);
        assert_eq!(references, vec![4, 7, 11]);
    }

    #[test]
    fn graph_ids_are_consecutive_and_checked() {
        let ids = GraphIds::allocate(40).expect("ids");
        assert_eq!(ids.all(), [40, 41, 42, 43]);
        assert_eq!(ids.last(), 43);
        assert!(GraphIds::allocate(u64::MAX - 2).is_none());
    }

    #[test]
    fn repoint_and_read_back_a_label_reference() {
        let mut drawable = Vec::new();
        litchi_iwa_common::wire::append_length_delimited_field(
            &mut drawable,
            Label::Caption.field(),
            &reference_payload(12),
            WireLimits::DEFAULT,
        )
        .expect("drawable");
        let mut message = Vec::new();
        litchi_iwa_common::wire::append_length_delimited_field(
            &mut message,
            1,
            &drawable,
            WireLimits::DEFAULT,
        )
        .expect("message");
        assert_eq!(
            label_reference(&message, Nesting::Drawable, Label::Caption),
            Ok(Some(12))
        );
        assert_eq!(
            label_reference(&message, Nesting::Drawable, Label::Title),
            Ok(None)
        );
        let moved = repoint(&message, Nesting::Drawable, Label::Caption, 300).expect("repoint");
        assert_eq!(
            label_reference(&moved, Nesting::Drawable, Label::Caption),
            Ok(Some(300))
        );
        assert_eq!(drawable_width(&moved, Nesting::Drawable), Ok(None));
    }
}
