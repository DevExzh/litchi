//! Theme layouts: discovery, selection, and materialization of a layout's
//! template slide into a fresh live slide (a port of the migration host's
//! `insert_slide`).

use std::collections::{HashMap, HashSet, VecDeque};

use litchi_core::Position;
use litchi_iwa_core::{ArchiveObject, RawMessage};

use super::graph::{self, MemberKind};
use super::{Error, Layout, LayoutSelector, Lift, Unsupported};
use crate::package::structure::metadata::ExternalReference;
use crate::package::structure::refs;
use crate::package::structure::wire::{self, Fields, Writer};
use crate::package::structure::workspace::Workspace;
use crate::package::{
    NOTE_MESSAGE_TYPE, PLACEHOLDER_MESSAGE_TYPE, Package, SHOW_MESSAGE_TYPE, SLIDE_MESSAGE_TYPE,
    SLIDE_NODE_MESSAGE_TYPE, STORAGE_MESSAGE_TYPE,
};

const THEME_MESSAGE_TYPE: u32 = 10;
const IMAGE_MESSAGE_TYPE: u32 = 3_005;
const MOVIE_MESSAGE_TYPE: u32 = 3_007;
const GUIDE_STORAGE_MESSAGE_TYPE: u32 = 3_047;
const SLIDE_NUMBER_ATTACHMENT_MESSAGE_TYPE: u32 = 2_043;
const IWORK_MESSAGE_VERSIONS: [u32; 3] = [1, 0, 5];
/// The preferred locator native Keynote gives live slide components.
const SLIDE_PREFERRED_LOCATOR: &str = "Slide";

const SHOW_THEME: u32 = 2;
const THEME_TEMPLATES: u32 = 2;
const THEME_DEFAULT_NODE: u32 = 5;
const THEME_DEFAULT_NODE_REFERENCE: u32 = 6;

const SLIDE_BUILDS: u32 = 2;
const SLIDE_DEPRECATED_BUILD_CHUNKS: u32 = 3;
const SLIDE_TITLE_PLACEHOLDER: u32 = 5;
const SLIDE_BODY_PLACEHOLDER: u32 = 6;
const SLIDE_OWNED_DRAWABLES: u32 = 7;
const SLIDE_NAME: u32 = 10;
const SLIDE_TEMPLATE: u32 = 17;
const SLIDE_STATIC_GUIDES: u32 = 18;
const SLIDE_IN_DOCUMENT: u32 = 19;
const SLIDE_NUMBER_PLACEHOLDER: u32 = 20;
const SLIDE_NOTE: u32 = 27;
const SLIDE_SAGE_TAGS: u32 = 28;
const SLIDE_BODY_PARAGRAPH_STYLES: u32 = 31;
const SLIDE_BODY_LIST_STYLES: u32 = 35;
const SLIDE_USER_GUIDES: u32 = 36;
const SLIDE_DRAWABLES_Z_ORDER: u32 = 42;
const SLIDE_BUILD_CHUNKS: u32 = 43;
const SLIDE_INFOS_USING_OBJECT_GEOMETRY: u32 = 44;
const SLIDE_INSTRUCTIONAL_TEXT_MAP: u32 = 45;
/// Layout-only singular fields a live slide does not carry.
const SLIDE_LAYOUT_ONLY_FIELDS: [u32; 21] = [
    10, // name
    11, // title placeholder geometry
    12, // title placeholder shape style index
    13, // title placeholder text style index
    14, // body placeholder geometry
    15, // body placeholder shape style index
    16, // body placeholder text style index
    21, // slide-number placeholder geometry
    22, // slide-number placeholder shape style index
    23, // slide-number placeholder text style index
    24, // title layout properties
    25, // body layout properties
    26, // slide-number layout properties
    29, // classic stylesheet record
    30, // object placeholder
    34, // deprecated object-placeholder export visibility
    37, // title thumbnail text
    38, // body thumbnail text
    39, // info using object-placeholder geometry
    40, // ... matches object-placeholder geometry
    41, // slide objects layer with template
];

const NOTE_STORAGE: u32 = 1;
const PLACEHOLDER_SHAPE: u32 = 1;
const SHAPE_OWNED_STORAGE: u32 = 4;
const STORAGE_TEXT: u32 = 3;
const STORAGE_ATTACHMENTS: u32 = 9;
const GUIDE_STORAGE_GUIDES: u32 = 1;
const DRAWABLE_SUPER: u32 = 1;
const DRAWABLE_PARENT: u32 = 2;
const IMAGE_FLAGS: u32 = 7;
const IMAGE_TEMPLATE_MEDIA_FLAG: u64 = 1 << 1;
const MOVIE_FLAGS: u32 = 13;
const MOVIE_MEDIA_PLACEHOLDER_FLAG: u64 = 1;
const MOVIE_IS_LIVE_VIDEO: u32 = 30;
const OBJECT_REPLACEMENT_CHARACTER: &str = "\u{fffc}";
const TEXTUAL_ATTACHMENT_PAGE_NUMBER: u64 = 0;

/// One theme layout's private native identity.
#[derive(Debug, Clone)]
pub(in crate::package) struct Record {
    node: u64,
    slide: u64,
    member: String,
    name: Option<String>,
}

impl Record {
    /// The layout's template component member.
    pub(in crate::package) fn member(&self) -> &str {
        &self.member
    }

    /// The layout's slide node.
    pub(in crate::package) const fn node(&self) -> u64 {
        self.node
    }

    /// The layout's template slide.
    pub(in crate::package) const fn slide(&self) -> u64 {
        self.slide
    }
}

/// The theme's layouts, read once per edit.
#[derive(Debug, Clone)]
pub(in crate::package) struct Catalog {
    records: Vec<Record>,
    default: Option<usize>,
}

fn unique_payload(package: &Package, identifier: u64, message_type: u32) -> Result<&[u8], Error> {
    let object = package.object(identifier).ok_or(Error::InvalidSource)?;
    let index = graph::message_index(object, message_type)?;
    Ok(&object.messages[index].data)
}

fn reference(fields: &Fields, number: u32) -> Result<Option<u64>, Error> {
    graph::reference_field(fields, number)
}

impl Catalog {
    /// Read the theme's layout list.
    pub(in crate::package) fn read(package: &Package) -> Result<Self, Error> {
        let limits = package
            .wire_limits()
            .map_err(|_error| Error::InvalidSource)?;
        let show_identifier = package
            .root_show_identifier()
            .map_err(|_error| Error::InvalidSource)?;
        let show = Fields::parse(
            unique_payload(package, show_identifier, SHOW_MESSAGE_TYPE)?,
            limits,
        )
        .lift()?;
        let theme = reference(&show, SHOW_THEME)?.ok_or(Error::InvalidSource)?;
        let theme =
            Fields::parse(unique_payload(package, theme, THEME_MESSAGE_TYPE)?, limits).lift()?;
        let nodes = graph::repeated_references(&theme, THEME_TEMPLATES)?;
        let default_node = match reference(&theme, THEME_DEFAULT_NODE_REFERENCE)? {
            Some(node) => Some(node),
            None => reference(&theme, THEME_DEFAULT_NODE)?,
        };
        let mut records = Vec::with_capacity(nodes.len());
        let mut seen = HashSet::with_capacity(nodes.len());
        for node in nodes {
            if node == 0 || !seen.insert(node) {
                return Err(Error::InvalidSource);
            }
            let fields = Fields::parse(
                unique_payload(package, node, SLIDE_NODE_MESSAGE_TYPE)?,
                limits,
            )
            .lift()?;
            let slide = reference(&fields, 2)?.ok_or(Error::InvalidSource)?;
            let member = package
                .structure_object_component(slide)
                .ok_or(Error::InvalidSource)?
                .to_owned();
            if graph::member_kind(&member) != MemberKind::TemplateSlide {
                return Err(Error::InvalidSource);
            }
            let slide_fields =
                Fields::parse(unique_payload(package, slide, SLIDE_MESSAGE_TYPE)?, limits)
                    .lift()?;
            let name = match slide_fields.unique(SLIDE_NAME).lift()? {
                Some(field) => Some(
                    std::str::from_utf8(field.message().lift()?)
                        .map_err(|_error| Error::InvalidSource)?
                        .to_owned(),
                ),
                None => None,
            };
            records.push(Record {
                node,
                slide,
                member,
                name,
            });
        }
        let default =
            default_node.and_then(|node| records.iter().position(|record| record.node == node));
        Ok(Self { records, default })
    }

    fn layout(&self, index: usize) -> Layout {
        Layout {
            name: self.records[index].name.clone(),
            position: Position::new(index),
        }
    }

    /// The public layout list.
    pub(in crate::package) fn layouts(&self) -> Box<[Layout]> {
        (0..self.records.len())
            .map(|index| self.layout(index))
            .collect()
    }

    /// The theme's default layout.
    pub(in crate::package) fn default_layout(&self) -> Option<Layout> {
        self.default.map(|index| self.layout(index))
    }

    /// Resolve a selector to a layout index.
    pub(in crate::package) fn select(&self, selector: LayoutSelector<'_>) -> Result<usize, Error> {
        match selector {
            LayoutSelector::Name(name) => {
                let mut matches = self
                    .records
                    .iter()
                    .enumerate()
                    .filter(|(_, record)| record.name.as_deref() == Some(name));
                let (index, _) = matches.next().ok_or(Error::Refused(
                    crate::slide::creation::Refusal::LayoutNameNotFound,
                ))?;
                if matches.next().is_some() {
                    return Err(Error::Refused(
                        crate::slide::creation::Refusal::AmbiguousLayout,
                    ));
                }
                Ok(index)
            },
            LayoutSelector::Position(position) => {
                if position.get() < self.records.len() {
                    Ok(position.get())
                } else {
                    Err(Error::Refused(
                        crate::slide::creation::Refusal::LayoutPositionNotFound { position },
                    ))
                }
            },
        }
    }

    /// The layout each slide of the show was made from, in show order:
    /// its template slide (`KN.SlideArchive.template_slide`) matched against
    /// this catalog; `None` for a slide without one or with a template the
    /// theme no longer lists.
    pub(in crate::package) fn assignments(
        &self,
        package: &Package,
    ) -> Result<Box<[Option<Layout>]>, Error> {
        let limits = package
            .wire_limits()
            .map_err(|_error| Error::InvalidSource)?;
        let nodes = package
            .private_slide_node_order()
            .map_err(|_error| Error::InvalidSource)?;
        let mut layouts = Vec::with_capacity(nodes.len());
        for node in &nodes {
            let fields = Fields::parse(
                unique_payload(package, *node, SLIDE_NODE_MESSAGE_TYPE)?,
                limits,
            )
            .lift()?;
            let slide = reference(&fields, 2)?.ok_or(Error::InvalidSource)?;
            let slide = Fields::parse(unique_payload(package, slide, SLIDE_MESSAGE_TYPE)?, limits)
                .lift()?;
            let template = reference(&slide, SLIDE_TEMPLATE)?;
            layouts.push(template.and_then(|template| {
                self.records
                    .iter()
                    .position(|record| record.slide == template)
                    .map(|index| self.layout(index))
            }));
        }
        Ok(layouts.into_boxed_slice())
    }

    pub(in crate::package) fn record(&self, index: usize) -> Result<&Record, Error> {
        self.records.get(index).ok_or(Error::InvalidSource)
    }
}

/// Objects reachable from `roots` inside `member`, following every object
/// reference (walked and header-listed). `stop` objects are selected but not
/// followed. Returns the selection in archive order.
pub(super) fn private_closure(
    workspace: &Workspace<'_>,
    member: &str,
    roots: impl IntoIterator<Item = u64>,
    stop: &[u64],
) -> Result<Vec<u64>, Error> {
    let internal = workspace.member_objects(member);
    let internal_set = internal.iter().copied().collect::<HashSet<_>>();
    let mut selected = stop.iter().copied().collect::<HashSet<_>>();
    let mut pending = roots.into_iter().collect::<VecDeque<_>>();
    while let Some(identifier) = pending.pop_front() {
        if !internal_set.contains(&identifier) {
            return Err(Error::InvalidSource);
        }
        if !selected.insert(identifier) {
            continue;
        }
        let object = workspace.require(identifier).lift()?;
        for reference in graph::all_object_references(object, workspace)? {
            if internal_set.contains(&reference) && !selected.contains(&reference) {
                pending.push_back(reference);
            }
        }
    }
    Ok(internal
        .into_iter()
        .filter(|identifier| selected.contains(identifier))
        .collect())
}

/// A live slide's speaker-notes graph reused for fresh slides.
pub(super) struct NoteSource {
    member: String,
    note: u64,
    objects: Vec<u64>,
}

impl NoteSource {
    /// Pick the smallest private notes graph among the base slides.
    pub(super) fn find(workspace: &Workspace<'_>, nodes: &[u64]) -> Result<Self, Error> {
        let mut best: Option<Self> = None;
        for node in nodes {
            let Ok(slide) = graph::node_slide(workspace, *node) else {
                continue;
            };
            let Some(member) = workspace.component_of(slide) else {
                continue;
            };
            let fields = Fields::parse(
                graph::payload(workspace, slide, SLIDE_MESSAGE_TYPE)?,
                workspace.wire,
            )
            .lift()?;
            let Some(note) = reference(&fields, SLIDE_NOTE)? else {
                continue;
            };
            if workspace.component_of(note).as_deref() != Some(member.as_str()) {
                continue;
            }
            let Ok(payload) = graph::payload(workspace, note, NOTE_MESSAGE_TYPE) else {
                continue;
            };
            let note_fields = Fields::parse(payload, workspace.wire).lift()?;
            let Some(storage) = reference(&note_fields, NOTE_STORAGE)? else {
                continue;
            };
            if workspace.component_of(storage).as_deref() != Some(member.as_str()) {
                continue;
            }
            let Ok(objects) = private_closure(workspace, &member, [note, storage], &[]) else {
                continue;
            };
            if objects.contains(&slide)
                || objects.iter().any(|object| {
                    workspace.require(*object).is_ok_and(|object| {
                        !object
                            .messages
                            .iter()
                            .zip(&object.archive_info.message_infos)
                            .all(|(message, info)| {
                                refs::admitted(message.type_)
                                    || (info.object_references.is_empty()
                                        && info.data_references.is_empty())
                            })
                    })
                })
            {
                continue;
            }
            if best
                .as_ref()
                .is_none_or(|current| objects.len() < current.objects.len())
            {
                best = Some(Self {
                    member,
                    note,
                    objects,
                });
            }
        }
        best.ok_or(Error::Refused(
            crate::slide::creation::Refusal::UnsupportedContent {
                reason: Unsupported::NotesTemplate,
            },
        ))
    }
}

/// How one owned layout drawable is materialized.
pub(super) enum Media {
    /// Not media: cloned as is.
    None,
    /// A template image: cloned with its template flag cleared.
    Image,
    /// A live-video or placeholder movie: cloned as is.
    Movie,
    /// A template-only movie: rendered by the layout, not cloned.
    TemplateOnly,
}

pub(super) fn classify(workspace: &Workspace<'_>, drawable: u64) -> Result<Media, Error> {
    let object = workspace.require(drawable).lift()?;
    let image = object
        .messages
        .iter()
        .any(|message| message.type_ == IMAGE_MESSAGE_TYPE);
    let movie = object
        .messages
        .iter()
        .find(|message| message.type_ == MOVIE_MESSAGE_TYPE);
    match (image, movie) {
        (false, None) => Ok(Media::None),
        (true, None) => Ok(Media::Image),
        (false, Some(movie)) => {
            let fields = Fields::parse(&movie.data, workspace.wire).lift()?;
            let live = match fields.unique(MOVIE_IS_LIVE_VIDEO).lift()? {
                Some(field) => field.varint().lift()? != 0,
                None => false,
            };
            let placeholder = match fields.unique(MOVIE_FLAGS).lift()? {
                Some(field) => field.varint().lift()? & MOVIE_MEDIA_PLACEHOLDER_FLAG != 0,
                None => false,
            };
            Ok(if live || placeholder {
                Media::Movie
            } else {
                Media::TemplateOnly
            })
        },
        (true, Some(_)) => Err(Error::Refused(
            crate::slide::creation::Refusal::UnsupportedContent {
                reason: Unsupported::LayoutMedia,
            },
        )),
    }
}

/// The drawable parent of a cloned media payload (`DrawableArchive.parent`).
pub(super) fn media_parent(
    payload: &[u8],
    workspace: &Workspace<'_>,
) -> Result<Option<u64>, Error> {
    let fields = Fields::parse(payload, workspace.wire).lift()?;
    let Some(drawable) = fields.unique(DRAWABLE_SUPER).lift()? else {
        return Ok(None);
    };
    let drawable = Fields::parse(drawable.message().lift()?, workspace.wire).lift()?;
    reference(&drawable, DRAWABLE_PARENT)
}

/// Materialize a fresh slide from `record` (and a notes graph cloned from
/// `notes`) into a new component; returns the new slide node.
pub(super) fn stage(
    workspace: &mut Workspace<'_>,
    record: &Record,
    notes: &NoteSource,
) -> Result<u64, Error> {
    let template = Fields::parse(
        graph::payload(workspace, record.slide, SLIDE_MESSAGE_TYPE)?,
        workspace.wire,
    )
    .lift()?;
    if template.all(SLIDE_BUILDS).next().is_some()
        || template.all(SLIDE_DEPRECATED_BUILD_CHUNKS).next().is_some()
        || template.all(SLIDE_BUILD_CHUNKS).next().is_some()
        || template.unique(SLIDE_NOTE).lift()?.is_some()
    {
        return Err(Error::Refused(
            crate::slide::creation::Refusal::UnsupportedContent {
                reason: Unsupported::LayoutState,
            },
        ));
    }
    let owned = graph::repeated_references(&template, SLIDE_OWNED_DRAWABLES)?;
    let mut images = Vec::new();
    let mut movies = Vec::new();
    let mut template_only = HashSet::new();
    for drawable in &owned {
        match classify(workspace, *drawable)? {
            Media::None => {},
            Media::Image => images.push(*drawable),
            Media::Movie => movies.push(*drawable),
            Media::TemplateOnly => {
                template_only.insert(*drawable);
            },
        }
    }
    let mut roots = owned
        .iter()
        .copied()
        .filter(|drawable| !template_only.contains(drawable))
        .collect::<Vec<_>>();
    for number in [
        SLIDE_TITLE_PLACEHOLDER,
        SLIDE_BODY_PLACEHOLDER,
        SLIDE_NUMBER_PLACEHOLDER,
        SLIDE_USER_GUIDES,
    ] {
        roots.extend(reference(&template, number)?);
    }
    let objects = private_closure(workspace, &record.member, roots, &[record.slide])?;
    if objects.iter().any(|object| template_only.contains(object)) {
        return Err(Error::Refused(
            crate::slide::creation::Refusal::UnsupportedContent {
                reason: Unsupported::LayoutMedia,
            },
        ));
    }

    let new_node = workspace.allocate().lift()?;
    let mut remap = graph::allocate_remap(workspace, &objects)?;
    let note_remap = graph::allocate_remap(workspace, &notes.objects)?;
    remap.extend(note_remap.iter().map(|(source, target)| (*source, *target)));
    let new_slide = *remap.get(&record.slide).ok_or(Error::InvalidSource)?;
    let new_note = *remap.get(&notes.note).ok_or(Error::InvalidSource)?;
    let locator = format!("Slide-{new_slide}");
    let new_member = format!("Index/{locator}.iwa");
    workspace.create_component(&new_member).lift()?;
    graph::clone_into(workspace, &new_member, &objects, &remap)?;
    graph::clone_into(workspace, &new_member, &notes.objects, &remap)?;

    materialize_slide(workspace, record.slide, new_slide, new_note, &remap)?;
    for image in &images {
        materialize_image(workspace, remap[image], new_slide)?;
    }
    for movie in &movies {
        let payload = graph::payload(workspace, remap[movie], MOVIE_MESSAGE_TYPE)?;
        if media_parent(payload, workspace)? != Some(new_slide) {
            return Err(Error::Refused(
                crate::slide::creation::Refusal::UnsupportedContent {
                    reason: Unsupported::LayoutMedia,
                },
            ));
        }
    }
    if let Some(guides) = reference(&template, SLIDE_USER_GUIDES)? {
        clear_guides(workspace, remap[&guides])?;
    }
    let attachment = match reference(&template, SLIDE_NUMBER_PLACEHOLDER)? {
        Some(placeholder) => prepare_slide_number(workspace, &new_member, remap[&placeholder])?,
        None => None,
    };
    graph::clone_node(workspace, record.node, new_node, record.slide, new_slide)?;
    register(
        workspace,
        record,
        notes,
        &new_member,
        &locator,
        new_slide,
        &remap,
        attachment,
    )?;
    graph::register_node(workspace, record.node, new_node, new_slide)?;
    Ok(new_node)
}

/// Turn the cloned template slide into a live slide (the migration host's
/// `materialize_slide_object`).
fn materialize_slide(
    workspace: &mut Workspace<'_>,
    template: u64,
    slide: u64,
    note: u64,
    remap: &HashMap<u64, u64>,
) -> Result<(), Error> {
    let cloned = remap.values().copied().collect::<HashSet<_>>();
    let mut fields = Fields::parse(
        graph::payload(workspace, slide, SLIDE_MESSAGE_TYPE)?,
        workspace.wire,
    )
    .lift()?;
    for number in [
        SLIDE_BUILDS,
        SLIDE_DEPRECATED_BUILD_CHUNKS,
        SLIDE_STATIC_GUIDES,
        SLIDE_BODY_PARAGRAPH_STYLES,
        SLIDE_BODY_LIST_STYLES,
        SLIDE_BUILD_CHUNKS,
        SLIDE_INFOS_USING_OBJECT_GEOMETRY,
    ] {
        fields.remove(number);
    }
    for number in SLIDE_LAYOUT_ONLY_FIELDS {
        fields.remove(number);
    }
    for number in [SLIDE_OWNED_DRAWABLES, SLIDE_DRAWABLES_Z_ORDER] {
        let retained = graph::repeated_references(&fields, number)?
            .into_iter()
            .filter(|drawable| cloned.contains(drawable))
            .map(wire::reference)
            .collect::<Vec<_>>();
        fields.replace_all(number, wire::LENGTH_DELIMITED, &retained);
    }
    let mut tags = Vec::new();
    for entry in fields.all(SLIDE_SAGE_TAGS) {
        let entry_fields = Fields::parse(entry.message().lift()?, workspace.wire).lift()?;
        if reference(&entry_fields, 2)?.is_some_and(|info| cloned.contains(&info)) {
            tags.push(entry.message().lift()?.to_vec());
        }
    }
    if fields.all(SLIDE_SAGE_TAGS).next().is_some() {
        fields.replace_all(SLIDE_SAGE_TAGS, wire::LENGTH_DELIMITED, &tags);
    }
    fields.set_message(SLIDE_INSTRUCTIONAL_TEXT_MAP, &[]);
    fields.set_message(SLIDE_TEMPLATE, &wire::reference(template));
    fields.set_varint(SLIDE_IN_DOCUMENT, 1);
    fields.set_message(SLIDE_NOTE, &wire::reference(note));
    graph::rewrite(workspace, slide, SLIDE_MESSAGE_TYPE, fields.encode(), None)
}

/// Clear a cloned layout image's template-media flag after checking that it
/// belongs to the new slide.
pub(super) fn materialize_image(
    workspace: &mut Workspace<'_>,
    image: u64,
    slide: u64,
) -> Result<(), Error> {
    let payload = graph::payload(workspace, image, IMAGE_MESSAGE_TYPE)?;
    if media_parent(payload, workspace)? != Some(slide) {
        return Err(Error::Refused(
            crate::slide::creation::Refusal::UnsupportedContent {
                reason: Unsupported::LayoutMedia,
            },
        ));
    }
    let mut fields = Fields::parse(payload, workspace.wire).lift()?;
    let Some(flags) = fields.unique(IMAGE_FLAGS).lift()? else {
        return Ok(());
    };
    let flags = flags.varint().lift()?;
    if flags & IMAGE_TEMPLATE_MEDIA_FLAG == 0 {
        return Ok(());
    }
    fields.set_varint(IMAGE_FLAGS, flags & !IMAGE_TEMPLATE_MEDIA_FLAG);
    graph::rewrite(workspace, image, IMAGE_MESSAGE_TYPE, fields.encode(), None)
}

/// Remove the user-defined guides a layout's guide storage carries.
fn clear_guides(workspace: &mut Workspace<'_>, storage: u64) -> Result<(), Error> {
    let mut fields = Fields::parse(
        graph::payload(workspace, storage, GUIDE_STORAGE_MESSAGE_TYPE)?,
        workspace.wire,
    )
    .lift()?;
    if fields.remove(GUIDE_STORAGE_GUIDES) == 0 {
        return Ok(());
    }
    graph::rewrite(
        workspace,
        storage,
        GUIDE_STORAGE_MESSAGE_TYPE,
        fields.encode(),
        None,
    )
}

/// Give a fresh slide-number placeholder storage its page-number attachment
/// (the migration host's `prepare_slide_number`). Returns the attachment.
fn prepare_slide_number(
    workspace: &mut Workspace<'_>,
    member: &str,
    placeholder: u64,
) -> Result<Option<u64>, Error> {
    let fields = Fields::parse(
        graph::payload(workspace, placeholder, PLACEHOLDER_MESSAGE_TYPE)?,
        workspace.wire,
    )
    .lift()?;
    let Some(shape) = fields.unique(PLACEHOLDER_SHAPE).lift()? else {
        return Ok(None);
    };
    let shape = Fields::parse(shape.message().lift()?, workspace.wire).lift()?;
    let Some(storage) = reference(&shape, SHAPE_OWNED_STORAGE)? else {
        return Ok(None);
    };
    if storage == 0 || workspace.component_of(storage).as_deref() != Some(member) {
        return Ok(None);
    }
    let mut storage_fields = Fields::parse(
        graph::payload(workspace, storage, STORAGE_MESSAGE_TYPE)?,
        workspace.wire,
    )
    .lift()?;
    if storage_fields.unique(STORAGE_ATTACHMENTS).lift()?.is_some() {
        return Ok(None);
    }
    let attachment = workspace.allocate().lift()?;
    let table = Writer::new()
        .message(1, &Writer::new().varint(1, 0).reference(2, attachment))
        .finish();
    storage_fields.replace_all(
        STORAGE_TEXT,
        wire::LENGTH_DELIMITED,
        &[OBJECT_REPLACEMENT_CHARACTER.as_bytes().to_vec()],
    );
    storage_fields.set_message(STORAGE_ATTACHMENTS, &table);
    let payload = Writer::new()
        .message(
            1,
            &Writer::new()
                .string(1, "")
                .varint(2, TEXTUAL_ATTACHMENT_PAGE_NUMBER),
        )
        .finish();
    let mut object = ArchiveObject::new(
        attachment,
        vec![RawMessage {
            type_: SLIDE_NUMBER_ATTACHMENT_MESSAGE_TYPE,
            data: payload,
        }],
        Default::default(),
    )
    .map_err(crate::package::structure::StructureError::from)
    .lift()?;
    object.archive_info.message_infos[0].versions = IWORK_MESSAGE_VERSIONS.to_vec().into();
    workspace.insert(member, object).lift()?;
    graph::rewrite(
        workspace,
        storage,
        STORAGE_MESSAGE_TYPE,
        storage_fields.encode(),
        None,
    )?;
    Ok(Some(attachment))
}

/// Register the fresh slide component: the layout component's records for
/// the cloned objects, the cloned notes graph, the slide-number attachment,
/// and external references recomputed from the component's objects.
#[allow(
    clippy::too_many_arguments,
    reason = "the registry step joins the layout, notes and attachment identities of one slide"
)]
fn register(
    workspace: &mut Workspace<'_>,
    record: &Record,
    notes: &NoteSource,
    member: &str,
    locator: &str,
    new_slide: u64,
    remap: &HashMap<u64, u64>,
    attachment: Option<u64>,
) -> Result<(), Error> {
    let layout = graph::registered(workspace, &record.member)?.clone();
    let note_component = graph::registered(workspace, &notes.member)?.clone();
    let mut component = layout
        .cloned(
            new_slide,
            locator,
            Some(SLIDE_PREFERRED_LOCATOR),
            &|object| remap.get(&object).copied(),
        )
        .lift()?;
    let note_remap = notes
        .objects
        .iter()
        .map(|object| (*object, remap[object]))
        .collect::<HashMap<_, _>>();
    graph::derive_uuids(workspace, &mut component, layout.object_uuids(), remap)?;
    graph::derive_uuids(
        workspace,
        &mut component,
        note_component.object_uuids(),
        &note_remap,
    )?;
    if let Some(attachment) = attachment {
        let uuid = workspace.clone_uuid(None, attachment).lift()?;
        component.add_object_uuid(attachment, uuid);
    }
    for object in &notes.objects {
        let target = remap[object];
        let source = workspace.require(*object).lift()?;
        let data = source
            .archive_info
            .message_infos
            .iter()
            .flat_map(|info| info.data_references.iter().copied())
            .collect::<Vec<_>>();
        for identifier in data {
            component.add_data_owner(identifier, target);
        }
    }

    let internal = workspace
        .member_objects(member)
        .into_iter()
        .collect::<HashSet<_>>();
    let mut needed = Vec::new();
    let template_link = ExternalReference {
        component: layout.identifier,
        object: None,
        is_weak: None,
    };
    needed.push(template_link);
    for object in workspace.member_objects(member) {
        let object = workspace.require(object).lift()?;
        for reference in graph::all_object_references(object, workspace)? {
            if internal.contains(&reference) || reference == record.slide {
                continue;
            }
            // A reference the source graph already left dangling stays as
            // it was; it needs no registry edge.
            let Some(target) = workspace.component_of(reference) else {
                continue;
            };
            if !matches!(
                graph::member_kind(&target),
                MemberKind::Document | MemberKind::Stylesheet
            ) {
                return Err(Error::Refused(
                    crate::slide::creation::Refusal::UnsupportedContent {
                        reason: Unsupported::ExternalContent,
                    },
                ));
            }
            let edge = ExternalReference {
                component: graph::registered(workspace, &target)?.identifier,
                object: Some(reference),
                is_weak: None,
            };
            if !needed.contains(&edge) {
                needed.push(edge);
            }
        }
    }
    let mut references = vec![template_link];
    references.extend(
        component
            .external_references()
            .iter()
            .filter(|reference| **reference != template_link && needed.contains(reference))
            .copied(),
    );
    for edge in needed {
        if !references.contains(&edge) {
            references.push(edge);
        }
    }
    component.set_external_references(references);
    workspace.registry.add_component(component).lift()
}
