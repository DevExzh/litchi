//! Package-level equation edits.
//!
//! [`Package::edit_equations`] stages equation additions, replacements and
//! removals over one immutable [`Package`]; [`Edit::commit`] publishes them
//! through one append-only user edit, and [`Package::apply_equations`]
//! replays a [`Patch`] onto the exact package it was made from. Dropping an
//! [`Edit`] discards it.
//!
//! A new equation is a MathType object (`Equation.DSMT4`, MTEF 5): its
//! `ExOleEmbedContainer` in the external-object list, its compressed
//! `ExOleObjStg` storage under a new persist identifier, a PNG preview in
//! the drawing group's picture store (BLIP in the `Pictures` stream), and a
//! picture-frame shape on the slide. A replacement keeps the object's kind
//! (an Equation Editor 3.0 object stays MTEF 3), its storage compression
//! and every other stream of its storage except cached presentations,
//! regenerates the preview and rescales the frame by the size change of
//! the rendered equation. A removal deletes the shape and, when nothing else
//! uses them, the object definition, its preview picture and the storage's
//! entry in the current persist directory.
//!
//! # Earlier bytes stay in the file
//!
//! Every PPT edit is an append-only incremental save (a new `UserEditAtom`
//! and persist directory after the existing records): the
//! records an edit replaces or drops are not erased. After a replacement or
//! a removal the old `ExOleObjStg` storage, and with it the old equation,
//! remains in the `PowerPoint Document` stream, and the persist directories
//! of earlier edits still map its persist identifier to it, so a reader
//! that walks the edit history (or the raw stream) can recover it. Do not
//! rely on these edits to remove confidential content from a file.

use std::collections::{BTreeMap, BTreeSet};
use std::sync::Arc;

use litchi_core::Position;
use litchi_core::edit::{Refusal, RefusalKind};
use litchi_equation::{Limits, MtefTarget, PngOptions};

use super::content::{self, Content};
use super::object;
use super::read::{
    self, DG_CONTAINER, FSP, PP_DRAWING, SP_CONTAINER, SPGR_CONTAINER, equation_definition,
};
use super::tree::{self, Node};
use super::{Bounds, Commit, Equation, Error, Patch, Result};
use crate::Package;
use crate::embedded::object::{
    Collection, ColorFollow, ContainerKind, Definition, DimensionPolicy, DrawAspect, Editor,
    EmbedPreferences, ExternalObject, Metadata, ObjectSubtype, ObjectType,
};
use crate::embedded::storage::{Kind as StorageKind, Storage};
use crate::package::{Error as PackageError, SourceLimits};
use crate::records::Record;
use crate::scoped::{self, Base};

/// Master units (576 per inch) per point.
const MASTER_UNITS_PER_POINT: f64 = 8.0;
/// The text size the size of an existing equation is measured at.
const MEASURE_POINTS: f64 = 12.0;
/// Pixels per CSS pixel of a preview, relative to the frame's size.
const PREVIEW_DENSITY: f64 = 2.0;
/// The smallest scale a preview is drawn at to fit the pixel ceiling.
const MIN_PREVIEW_SCALE: f64 = 1.0 / 64.0;
/// The `EQNOLEFILEHDR` length.
const OLE_HEADER_LEN: usize = 28;

/// The equation an add or a replacement writes.
#[derive(Clone, Copy, Debug)]
pub enum Source<'a> {
    /// LaTeX math, converted to MTEF.
    Latex(&'a str),
    /// An OMML equation (`m:oMath`, `m:oMathPara` or a fragment of them),
    /// converted to MTEF.
    Omml(&'a str),
    /// Raw MTEF: a complete `Equation Native` stream, or bare MTEF to which
    /// the 28-byte OLE header is added. A new object takes the kind of its
    /// MTEF version (3: `Equation.3`, 5: `Equation.DSMT4`); a replacement
    /// must have the version of the object it replaces.
    Mtef(&'a [u8]),
}

/// What an add or a replacement wrote.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Written {
    shape_id: u32,
    losses: Vec<String>,
}

impl Written {
    /// The shape identifier of the equation's frame.
    #[must_use]
    pub const fn shape_id(&self) -> u32 {
        self.shape_id
    }

    /// One message per piece of the source MTEF cannot hold; empty when
    /// nothing was lost (always empty for a [`Source::Mtef`]).
    #[must_use]
    pub fn losses(&self) -> &[String] {
        &self.losses
    }
}

/// A staged edit of the equations of one immutable [`Package`].
pub struct Edit {
    base: Base,
    state: State,
}

impl std::fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("changed", &self.state.changed)
            .finish_non_exhaustive()
    }
}

/// The staged records of an equation edit.
struct State {
    /// The record publisher; operations only read it until `finish`.
    editor: Arc<Editor>,
    source_limits: SourceLimits,
    /// The persist identifier of each slide, by position.
    slides: Vec<u32>,
    document_before: Vec<u8>,
    document: Node,
    collection: Collection,
    collection_changed: bool,
    /// Slides opened by this edit: their source bytes and expanded tree.
    edited: BTreeMap<usize, (Vec<u8>, Node)>,
    pictures_before: Option<Vec<u8>>,
    pictures: Vec<u8>,
    /// Storage records written (`Some`) or removed (`None`).
    storages: BTreeMap<u32, Option<Vec<u8>>>,
    /// Persist identifiers this edit allocated.
    inserted: BTreeSet<u32>,
    changed: bool,
}

/// One equation frame located on a slide.
struct Located {
    /// Child indexes from the patriarch group down to the frame.
    path: Vec<usize>,
    object_id: u32,
    persist_id: u32,
    bounds: Option<Bounds>,
    picture: Option<u32>,
    program: Option<String>,
}

fn refused(kind: RefusalKind, operation: &'static str, reason: &'static str) -> Error {
    Error::Refused(Refusal::new(kind, operation, reason))
}

impl Edit {
    fn new(package: &Package) -> Result<Self> {
        let base = Base::capture(package)?;
        let presentation = package.presentation()?;
        let source_limits = presentation.source_limits();
        let collection = presentation
            .ole_objects()?
            .unwrap_or_else(Collection::empty);
        let slides = presentation
            .slide_directory()
            .entries()
            .iter()
            .map(crate::SlideDirectoryEntry::persist_id)
            .collect();
        let editor =
            Editor::open_records_admitting_signature(base.bytes()?.to_vec(), source_limits)
                .map_err(scoped::publication("edit_equations"))?;
        let document_before = editor.persisted_record(editor.document_persist_id())?;
        let document = tree::parse(&document_before)?;
        let pictures_before = editor.stream_by_name("Pictures")?.map(<[u8]>::to_vec);
        Ok(Self {
            base,
            state: State {
                editor: Arc::new(editor),
                source_limits,
                slides,
                document_before,
                document,
                collection,
                collection_changed: false,
                edited: BTreeMap::new(),
                pictures: pictures_before.clone().unwrap_or_default(),
                pictures_before,
                storages: BTreeMap::new(),
                inserted: BTreeSet::new(),
                changed: false,
            },
        })
    }

    /// The staged equations of every slide, as
    /// [`Presentation::equations`](crate::Presentation::equations) lists
    /// them.
    ///
    /// # Errors
    ///
    /// Returns an error when a staged slide or the object list cannot be
    /// read; an unreadable storage makes only its equation's stream
    /// [`Unreadable`](super::Unreadable). Each storage is read once, within
    /// the source's aggregate payload limit.
    pub fn equations(&self) -> Result<Vec<Equation>> {
        let state = &self.state;
        let mut out = Vec::new();
        let mut read: BTreeMap<u32, (bool, Content)> = BTreeMap::new();
        let mut remaining = state.source_limits.max_aggregate_payload_bytes();
        for (slide, persist_id) in state.slides.iter().enumerate() {
            let record = match state.edited.get(&slide) {
                Some((_, node)) => node.encode()?,
                None => state.editor.persisted_record(*persist_id)?,
            };
            out.extend(read::slide_equations(
                slide,
                &record,
                &state.collection,
                &mut |persist_id| {
                    read.entry(persist_id)
                        .or_insert_with(|| {
                            let storage = state.storage(persist_id);
                            let compressed = matches!(
                                &storage,
                                Ok(Some(stored)) if stored.compression().is_zlib()
                            );
                            let native = if remaining == 0 {
                                Err(content::budget_reached())
                            } else {
                                let (native, spent) = content::load(storage, remaining);
                                remaining = remaining.saturating_sub(spent);
                                native
                            };
                            (compressed, Content::Loaded(native))
                        })
                        .clone()
                },
            )?);
        }
        Ok(out)
    }

    /// Stage a new equation on the slide at `slide`, its frame's top-left
    /// corner at (`left`, `top`) in master units (576 per inch) and its
    /// size that of the equation set at `font_size_pt` points. The object is
    /// a MathType object unless `source` is MTEF 3.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Equation`] when the source does not convert or
    /// render within `limits`, [`Error::Refused`] when the slide does not
    /// exist, has no drawing or no free shape identifier, or the frame does
    /// not fit the slide coordinates; the edit is then unchanged.
    pub fn add(
        &mut self,
        slide: Position,
        left: i32,
        top: i32,
        font_size_pt: f64,
        source: Source<'_>,
        limits: &Limits,
    ) -> Result<Written> {
        const OPERATION: &str = "equation.add";
        if slide.get() >= self.state.slides.len() {
            return Err(refused(
                RefusalKind::NotFound,
                OPERATION,
                "no slide has this position",
            ));
        }
        let (native, target, losses) = equation_native(source, None, limits)?;
        let options = PngOptions::new(font_size_pt)?.with_scale(PREVIEW_DENSITY)?;
        let preview = render_fitting(&native, options, limits)?;
        let (width, height) = preview.size_pt();
        let bounds = frame_bounds(left, top, width, height).ok_or_else(|| {
            refused(
                RefusalKind::InvalidValue,
                OPERATION,
                "the equation frame does not fit the slide coordinates",
            )
        })?;
        let storage = object::storage_record(object::new_storage(&target, &native)?, true)?;

        let mut state = self.state.clone_for_staging();
        let shape_id = state.allocate_shape(slide.get())?.ok_or_else(|| {
            refused(
                RefusalKind::UnsupportedTarget,
                OPERATION,
                "the slide has no drawing with a free shape identifier",
            )
        })?;
        let persist_id = state.next_persist_id()?;
        let object_id = state
            .collection
            .id_seed
            .checked_add(1)
            .filter(|id| i32::try_from(*id).is_ok())
            .ok_or_else(|| {
                refused(
                    RefusalKind::UnsupportedTarget,
                    OPERATION,
                    "the external-object list has no free identifier",
                )
            })?;
        state.collection.id_seed = object_id;
        state.collection.add(ExternalObject::Object(Definition {
            kind: ContainerKind::Embedded(EmbedPreferences {
                color_follow: ColorFollow::None,
                // As PowerPoint writes an equation's `ExOleEmbedAtom`.
                cannot_lock_server: true,
                dimension_policy: DimensionPolicy::Send,
                is_word_table: false,
                unused: 0,
            }),
            object: Metadata {
                draw_aspect: DrawAspect::Content,
                object_type: ObjectType::Embedded,
                id: object_id,
                subtype: ObjectSubtype::Equation,
                persist_id,
                unused: [0; 4],
            },
            menu_name: Some("Equation".to_owned()),
            program_id: Some(target.prog_id().to_owned()),
            clipboard_name: Some(target.user_type().to_owned()),
            metafile: None,
        }))?;
        state.collection_changed = true;
        state.storages.insert(persist_id, Some(storage));
        state.inserted.insert(persist_id);
        let picture =
            object::add_picture(&mut state.document, &mut state.pictures, preview.bytes())?;
        let shape = object::frame_shape(shape_id, object_id, picture, bounds)?;
        let slide_node = state.slide(slide.get())?;
        patriarch_mut(slide_node)?
            .children_mut()
            .ok_or_else(not_container)?
            .push(shape);
        state.changed = true;
        self.state = state;
        Ok(Written { shape_id, losses })
    }

    /// Stage a replacement of the equation whose frame has `shape_id` on the
    /// slide at `slide`. The object keeps its kind; the frame keeps its
    /// top-left corner and is scaled by the size change of the rendered
    /// equation. The previous storage stays in the file, reachable through
    /// earlier persist directories (see the module documentation).
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] with [`RefusalKind::NotFound`] when the
    /// slide has no such equation, [`Error::Equation`] when the source does
    /// not convert or render within `limits`, [`Error::Refused`] with
    /// [`RefusalKind::InvalidValue`] for raw MTEF of another version, and
    /// [`Error::Refused`] with [`RefusalKind::UnsupportedContent`] when the
    /// object's storage cannot be read; the edit is then unchanged.
    pub fn set(
        &mut self,
        slide: Position,
        shape_id: u32,
        source: Source<'_>,
        limits: &Limits,
    ) -> Result<Written> {
        const OPERATION: &str = "equation.set";
        let mut state = self.state.clone_for_staging();
        let located = state.locate(slide.get(), shape_id, OPERATION)?;
        let record = state.storage_record(located.persist_id)?;
        let stored = parse_storage(&record, &state.source_limits)?;
        let compound = stored
            .decompressed_bytes(read::MAX_STORAGE_BYTES)
            .map_err(|_error| {
                refused(
                    RefusalKind::UnsupportedContent,
                    OPERATION,
                    "the equation object's storage cannot be read",
                )
            })?;
        let old_native = content::load(Ok(Some(stored.clone())), read::MAX_STORAGE_BYTES)
            .0
            .map_err(|_unreadable| {
                refused(
                    RefusalKind::UnsupportedContent,
                    OPERATION,
                    "the equation object's storage cannot be read",
                )
            })?;
        // The existing equation's MTEF header names its kind; an equation the
        // reader cannot parse keeps the kind its ProgID names.
        let kind = litchi_equation::read_mtef(&old_native, limits)
            .ok()
            .and_then(|read| read.info().cloned())
            .map_or_else(
                || {
                    if located.program.as_deref() == Some("Equation.3") {
                        MtefTarget::EquationEditor3
                    } else {
                        MtefTarget::MathType
                    }
                },
                MtefTarget::SameAs,
            );
        let (native, _, losses) = equation_native(source, Some(kind), limits)?;
        let storage = object::storage_record(
            object::replace_native(&compound, &native)?,
            stored.compression().is_zlib(),
        )?;

        // The frame scales with the rendered equation.
        let measure = PngOptions::new(MEASURE_POINTS)?;
        let new_size = render_fitting(&native, measure, limits)?.size_pt();
        let old_size = render_fitting(&old_native, measure, limits)
            .ok()
            .map(|png| png.size_pt());
        let bounds = located
            .bounds
            .map(|bounds| rescaled(bounds, old_size, new_size))
            .transpose()
            .map_err(|()| {
                refused(
                    RefusalKind::InvalidValue,
                    OPERATION,
                    "the rescaled equation frame does not fit the slide coordinates",
                )
            })?;
        let zoom = bounds.map_or(1.0, |bounds| {
            (bounds.height() as f64 / MASTER_UNITS_PER_POINT / new_size.1.max(1.0)).clamp(0.1, 8.0)
        });
        let options = measure.with_scale((PREVIEW_DENSITY * zoom).clamp(1.0, 16.0))?;
        let preview = render_fitting(&native, options, limits)?;

        if let Some(picture) = located.picture {
            object::release_picture(&mut state.document, &mut state.pictures, picture)?;
        }
        let picture =
            object::add_picture(&mut state.document, &mut state.pictures, preview.bytes())?;
        let frame = state.frame_mut(slide.get(), &located.path)?;
        object::set_picture(frame, picture)?;
        if let Some(bounds) = bounds {
            object::set_bounds(frame, bounds)?;
        }
        state.storages.insert(located.persist_id, Some(storage));
        state.changed = true;
        self.state = state;
        Ok(Written { shape_id, losses })
    }

    /// Stage the removal of the equation whose frame has `shape_id` on the
    /// slide at `slide`: the frame and, when no other frame or object uses
    /// them, the object definition, its preview picture and its storage's
    /// persist-directory entry. The storage's bytes stay in the file,
    /// reachable through earlier persist directories (see the module
    /// documentation).
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] with [`RefusalKind::NotFound`] when the
    /// slide has no such equation, and with
    /// [`RefusalKind::UnsupportedContent`] (the reason names the record)
    /// when an animation, a build list or a connector, arc or callout rule
    /// of the slide still names the frame; the edit is then unchanged.
    pub fn remove(&mut self, slide: Position, shape_id: u32) -> Result<()> {
        const OPERATION: &str = "equation.remove";
        let mut state = self.state.clone_for_staging();
        let located = state.locate(slide.get(), shape_id, OPERATION)?;
        let slide_node = state.slide(slide.get())?;
        // The frame cannot go while the slide still names it: the crate
        // does not rewrite timing, build-list or connector records.
        if let Some(reason) = shape_reference(&slide_node.encode()?, shape_id) {
            return Err(refused(RefusalKind::UnsupportedContent, OPERATION, reason));
        }
        let (parent_path, index) = located
            .path
            .split_last()
            .map(|(last, parent)| (parent.to_vec(), *last))
            .ok_or_else(not_container)?;
        let mut parent = patriarch_mut(slide_node)?;
        for step in &parent_path {
            parent = parent
                .children_mut()
                .and_then(|children| children.get_mut(*step))
                .ok_or_else(not_container)?;
        }
        parent
            .children_mut()
            .ok_or_else(not_container)?
            .remove(index);
        adjust_shape_count(slide_node, false)?;
        state.unregister_shape()?;
        if let Some(picture) = located.picture {
            object::release_picture(&mut state.document, &mut state.pictures, picture)?;
        }
        if !state.object_shown(located.object_id)? {
            state.collection.remove(located.object_id)?;
            state.collection_changed = true;
            let shared = state
                .collection
                .objects
                .iter()
                .any(|object| object.persist_id() == located.persist_id);
            if !shared {
                state.storages.insert(located.persist_id, None);
            }
        }
        state.changed = true;
        self.state = state;
        Ok(())
    }

    /// Publish the staged equations as a new package. An edit with no net
    /// change publishes the source package unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Signed`] for a changed write to a signed package and
    /// [`Error::Package`] when the result cannot be republished.
    pub fn commit(self) -> Result<Commit> {
        let Self { base, state } = self;
        if !state.changed {
            return Ok(super::published(base.unchanged()?));
        }
        if base.is_signed()? {
            return Err(Error::Signed);
        }
        let after = state.finish()?;
        Ok(super::published(base.publish(after)?))
    }
}

impl State {
    /// A copy to stage one operation on, so a failed operation leaves the
    /// edit unchanged.
    fn clone_for_staging(&self) -> Self {
        Self {
            editor: Arc::clone(&self.editor),
            source_limits: self.source_limits,
            slides: self.slides.clone(),
            document_before: self.document_before.clone(),
            document: self.document.clone(),
            collection: self.collection.clone(),
            collection_changed: self.collection_changed,
            edited: self.edited.clone(),
            pictures_before: self.pictures_before.clone(),
            pictures: self.pictures.clone(),
            storages: self.storages.clone(),
            inserted: self.inserted.clone(),
            changed: self.changed,
        }
    }

    /// The expanded tree of the slide at `position`, opened on first use.
    fn slide(&mut self, position: usize) -> Result<&mut Node> {
        if !self.edited.contains_key(&position) {
            let persist_id = *self.slides.get(position).ok_or_else(|| {
                refused(
                    RefusalKind::NotFound,
                    "equation.edit",
                    "no slide has this position",
                )
            })?;
            let before = self.editor.persisted_record(persist_id)?;
            let mut node = tree::parse(&before)?;
            node.expand_all(0)?;
            self.edited.insert(position, (before, node));
        }
        self.edited
            .get_mut(&position)
            .map(|(_, node)| node)
            .ok_or_else(not_container)
    }

    /// The live or staged storage of `persist_id`.
    fn storage(&self, persist_id: u32) -> crate::package::Result<Option<Storage>> {
        match self.storages.get(&persist_id) {
            Some(None) => Ok(None),
            Some(Some(record)) => parse_storage(record, &self.source_limits).map(Some),
            None => {
                let record = self.editor.persisted_record(persist_id)?;
                parse_storage(&record, &self.source_limits).map(Some)
            },
        }
    }

    fn storage_record(&self, persist_id: u32) -> Result<Vec<u8>> {
        match self.storages.get(&persist_id) {
            Some(Some(record)) => Ok(record.clone()),
            Some(None) => Err(not_container()),
            None => Ok(self.editor.persisted_record(persist_id)?),
        }
    }

    fn next_persist_id(&self) -> Result<u32> {
        let staged = self.inserted.iter().next_back().copied().unwrap_or(0);
        Ok(self.editor.next_persist_id()?.max(staged.saturating_add(1)))
    }

    /// Find the equation frame `shape_id` on the slide at `position`.
    fn locate(
        &mut self,
        position: usize,
        shape_id: u32,
        operation: &'static str,
    ) -> Result<Located> {
        let not_found = || {
            refused(
                RefusalKind::NotFound,
                operation,
                "the slide has no equation with this shape identifier",
            )
        };
        if position >= self.slides.len() {
            return Err(not_found());
        }
        let collection = self.collection.clone();
        let slide = self.slide(position)?;
        let group = read::patriarch(slide).ok_or_else(not_found)?;
        let path = find_shape(group, shape_id, 0).ok_or_else(not_found)?;
        let mut node = group;
        for step in &path {
            node = &node.children()[*step];
        }
        let shape = read::object_shape(node, path.len() > 1).ok_or_else(not_found)?;
        let (persist_id, program) =
            equation_definition(&collection, shape.object_id).ok_or_else(not_found)?;
        Ok(Located {
            path,
            object_id: shape.object_id,
            persist_id,
            bounds: shape.bounds,
            picture: shape.picture,
            program: program.map(str::to_owned),
        })
    }

    fn frame_mut(&mut self, position: usize, path: &[usize]) -> Result<&mut Node> {
        let mut node = patriarch_mut(self.slide(position)?)?;
        for step in path {
            node = node
                .children_mut()
                .and_then(|children| children.get_mut(*step))
                .ok_or_else(not_container)?;
        }
        Ok(node)
    }

    /// Whether any slide still shows external object `object_id`.
    fn object_shown(&mut self, object_id: u32) -> Result<bool> {
        for position in 0..self.slides.len() {
            let shapes = match self.edited.get(&position) {
                Some((_, node)) => read::object_shapes(&node.encode()?)?,
                None => read::object_shapes(&self.editor.persisted_record(self.slides[position])?)?,
            };
            if shapes.iter().any(|shape| shape.object_id == object_id) {
                return Ok(true);
            }
        }
        Ok(false)
    }

    /// Allocate the shape identifier of a new shape on the slide at
    /// `position` and count the shape in its drawing (`OfficeArtFDG`) and
    /// the drawing group (`OfficeArtFDGG`). The identifier is unused on
    /// every slide. It follows every one the drawing used when that cluster
    /// belongs to the drawing; otherwise it is the next free one of a
    /// cluster the drawing owns, or the first of a new cluster (MS-ODRAW
    /// 2.2.46-2.2.48), as PowerPoint allocates them. `None` for a slide
    /// without a drawing or when no identifier is left.
    fn allocate_shape(&mut self, position: usize) -> Result<Option<u32>> {
        const OPERATION: &str = "equation.add";
        // Every drawing's shapes: the other slides (as staged), the masters
        // and the notes pages.
        let mut everywhere = Vec::new();
        for (other, (_, node)) in &self.edited {
            if *other != position {
                shape_ids(node, &mut everywhere, 0);
            }
        }
        let edited = self
            .edited
            .keys()
            .filter_map(|slide| self.slides.get(*slide).copied())
            .collect::<BTreeSet<_>>();
        for persist_id in self.editor.persist_ids() {
            if edited.contains(&persist_id) {
                continue;
            }
            let Ok(record) = self.editor.source_record(persist_id) else {
                continue;
            };
            let kind = record
                .get(2..4)
                .map(|kind| u16::from_le_bytes([kind[0], kind[1]]));
            if !matches!(kind, Some(SLIDE | MAIN_MASTER | NOTES)) {
                continue;
            }
            let mut node = tree::parse(record)?;
            node.expand_all(0)?;
            shape_ids(&node, &mut everywhere, 0);
        }
        let slide = self.slide(position)?;
        let Some(container) = slide
            .child_mut(PP_DRAWING)?
            .and_then(|drawing| drawing.child_mut(DG_CONTAINER).ok().flatten())
        else {
            return Ok(None);
        };
        let Some(fdg) = container
            .child(FDG)
            .filter(|fdg| fdg.payload().is_some_and(|payload| payload.len() >= 8))
        else {
            return Ok(None);
        };
        let drawing_id = u32::from(fdg.instance());
        let current = fdg.payload().map_or(0, |payload| {
            u32::from_le_bytes([payload[4], payload[5], payload[6], payload[7]])
        });
        let mut own = Vec::new();
        shape_ids(container, &mut own, 0);
        everywhere.extend_from_slice(&own);
        let taken_anywhere = |spid: u32| everywhere.contains(&spid);

        let group = object::drawing_group(&mut self.document)?;
        let fdgg = group
            .child_mut(FDGG)?
            .and_then(Node::payload_mut)
            .filter(|payload| payload.len() >= 16 && (payload.len() - 16) % 8 == 0)
            .ok_or_else(|| PackageError::Corrupted("the PPT drawing group has no FDGG".into()))?;
        let field = |bytes: &[u8], at: usize| {
            u32::from_le_bytes([bytes[at], bytes[at + 1], bytes[at + 2], bytes[at + 3]])
        };
        let clusters = (fdgg.len() - 16) / 8;
        // The drawing group's counts come from the file: refuse a table
        // whose values no valid drawing group holds rather than compute
        // with them.
        let invalid = field(fdgg, 0) > MAX_SPID
            || usize::try_from(field(fdgg, 4)).ok() != clusters.checked_add(1)
            || (0..clusters).any(|index| field(fdgg, 16 + index * 8 + 4) > 1024);
        if invalid {
            return Err(refused(
                RefusalKind::UnsupportedContent,
                OPERATION,
                "the drawing group's shape-identifier table holds out-of-range values",
            ));
        }
        // The cluster (1-based) of a shape identifier.
        let entry = |spid: u32| (spid >> 10) as usize;
        let owned = |bytes: &[u8], cluster: usize| {
            cluster >= 1
                && cluster <= clusters
                && field(bytes, 16 + (cluster - 1) * 8) == drawing_id
        };
        let after_own = own
            .iter()
            .copied()
            .max()
            .unwrap_or(0)
            .max(current)
            .checked_add(1)
            .filter(|next| *next & 0x3FF != 0 && owned(fdgg, entry(*next)))
            .and_then(|next| (next..=(next | 0x3FF)).find(|spid| !taken_anywhere(*spid)));
        let in_owned = || {
            (1..=clusters)
                .filter(|cluster| owned(fdgg, *cluster))
                .find_map(|cluster| {
                    let taken = field(fdgg, 16 + (cluster - 1) * 8 + 4);
                    let base = u32::try_from(cluster).ok()?.checked_mul(1024)?;
                    (taken..1024)
                        .filter_map(|offset| base.checked_add(offset))
                        .find(|spid| !taken_anywhere(*spid))
                })
        };
        let shape_id = if let Some(spid) = after_own.or_else(in_owned) {
            spid
        } else {
            // A new cluster for this drawing.
            let cluster = clusters + 1;
            let Some(base) = u32::try_from(cluster)
                .ok()
                .and_then(|cluster| cluster.checked_mul(1024))
                .filter(|base| *base < MAX_SPID)
            else {
                return Ok(None);
            };
            let Some(spid) = (base..base + 1024).find(|spid| !taken_anywhere(*spid)) else {
                return Ok(None);
            };
            fdgg.extend_from_slice(&drawing_id.to_le_bytes());
            fdgg.extend_from_slice(&0u32.to_le_bytes());
            let cidcl = u32::try_from(cluster + 1).unwrap_or(u32::MAX);
            fdgg[4..8].copy_from_slice(&cidcl.to_le_bytes());
            spid
        };
        let at = 16 + (entry(shape_id) - 1) * 8;
        // At most 1024: the offset is below 1024 and the stored value was
        // checked above.
        let taken = field(fdgg, at + 4).max((shape_id & 0x3FF) + 1);
        fdgg[at + 4..at + 8].copy_from_slice(&taken.to_le_bytes());
        let saved = field(fdgg, 8).saturating_add(1);
        fdgg[8..12].copy_from_slice(&saved.to_le_bytes());
        let maximum = field(fdgg, 0)
            .max((shape_id & !0x3FF).saturating_add(taken))
            .min(MAX_SPID);
        fdgg[..4].copy_from_slice(&maximum.to_le_bytes());

        adjust_shape_count(self.slide(position)?, true)?;
        let fdg = self
            .slide(position)?
            .child_mut(PP_DRAWING)?
            .and_then(|drawing| drawing.child_mut(DG_CONTAINER).ok().flatten())
            .and_then(|container| container.child_mut(FDG).ok().flatten())
            .and_then(Node::payload_mut)
            .ok_or_else(not_container)?;
        fdg[4..8].copy_from_slice(&current.max(shape_id).to_le_bytes());
        Ok(Some(shape_id))
    }

    /// Count one shape fewer in the drawing group.
    fn unregister_shape(&mut self) -> Result<()> {
        let group = object::drawing_group(&mut self.document)?;
        if let Some(fdgg) = group.child_mut(FDGG)?.and_then(Node::payload_mut)
            && fdgg.len() >= 16
        {
            let saved =
                u32::from_le_bytes([fdgg[8], fdgg[9], fdgg[10], fdgg[11]]).saturating_sub(1);
            fdgg[8..12].copy_from_slice(&saved.to_le_bytes());
        }
        Ok(())
    }

    /// The bytes of the published package.
    fn finish(self) -> Result<Vec<u8>> {
        let Self {
            editor,
            source_limits,
            slides,
            document_before,
            document,
            collection,
            collection_changed,
            edited,
            pictures_before,
            pictures,
            storages,
            inserted,
            changed: _,
        } = self;
        let mut editor = Arc::try_unwrap(editor).unwrap_or_else(|shared| (*shared).clone());
        let mut document = document.encode()?;
        if collection_changed {
            let list = collection.to_record_bytes(&source_limits)?;
            document =
                crate::embedded::object::editor::replace_or_insert_object_list(&document, &list)?;
        }
        if document != document_before {
            editor.replace_persisted_record(editor.document_persist_id(), document)?;
        }
        for (position, (before, node)) in &edited {
            let after = node.encode()?;
            if &after != before {
                editor.replace_persisted_record(slides[*position], after)?;
            }
        }
        for (persist_id, record) in storages {
            match record {
                Some(record) if inserted.contains(&persist_id) => {
                    editor.insert_persisted_record(persist_id, record)?;
                },
                Some(record) => editor.replace_persisted_record(persist_id, record)?,
                None if inserted.contains(&persist_id) => {},
                None => editor.remove_persisted_record(persist_id)?,
            }
        }
        let pictures_changed = match &pictures_before {
            Some(before) => *before != pictures,
            None => !pictures.is_empty(),
        };
        if pictures_changed {
            editor.replace_stream_by_name("Pictures", pictures)?;
        }
        Ok(editor.finish()?)
    }
}

/// `SlideContainer`, `MainMasterContainer` and `NotesContainer`.
const SLIDE: u16 = 0x03EE;
const MAIN_MASTER: u16 = 0x03F8;
const NOTES: u16 = 0x03F0;
const FDGG: u16 = 0xF006;
const FDG: u16 = 0xF008;
/// The largest valid `spidMax` (MS-ODRAW 2.2.48).
const MAX_SPID: u32 = 0x03FF_D7FF;

fn not_container() -> Error {
    Error::Package(PackageError::Corrupted(
        "a PPT drawing record has an unexpected shape".into(),
    ))
}

fn parse_storage(record: &[u8], limits: &SourceLimits) -> crate::package::Result<Storage> {
    let (record, _) = Record::parse(record, 0, limits)?;
    Storage::parse_as(&record, StorageKind::OleObject)
}

fn patriarch_mut(slide: &mut Node) -> Result<&mut Node> {
    slide
        .child_mut(PP_DRAWING)?
        .ok_or_else(not_container)?
        .child_mut(DG_CONTAINER)?
        .ok_or_else(not_container)?
        .child_mut(SPGR_CONTAINER)?
        .ok_or_else(not_container)
}

/// `VisualShapeAtom` ([MS-PPT] 2.8.70).
const VISUAL_SHAPE_ATOM: u16 = 0x2AFB;
/// `BuildAtom` ([MS-PPT] 2.8.6), the head of every build-list entry.
const BUILD_ATOM: u16 = 0x2B03;
/// `BinaryTagDataBlob` ([MS-PPT] 2.4.22.3): a record sequence in an atom.
const BINARY_TAG_DATA_BLOB: u16 = 0x138B;
/// `OfficeArtFConnectorRule` ([MS-ODRAW] 2.2.36).
const CONNECTOR_RULE: u16 = 0xF012;
/// `OfficeArtFArcRule` ([MS-ODRAW] 2.2.38).
const ARC_RULE: u16 = 0xF014;
/// `OfficeArtFCalloutRule` ([MS-ODRAW] 2.2.39).
const CALLOUT_RULE: u16 = 0xF017;

/// What else in a slide record names shape `shape_id`: an animation target
/// (`VisualShapeAtom` of a shape), a build-list entry (`BuildAtom`), or a
/// connector, arc or callout rule. The scan enters containers and the
/// record sequences of programmable-tag data blobs.
pub(super) fn shape_reference(record: &[u8], shape_id: u32) -> Option<&'static str> {
    fn scan(bytes: &[u8], shape_id: u32, depth: usize) -> Option<&'static str> {
        if depth > 64 {
            return None;
        }
        let at = |payload: &[u8], offset: usize| -> Option<u32> {
            Some(u32::from_le_bytes(
                payload.get(offset..offset + 4)?.try_into().ok()?,
            ))
        };
        let mut offset = 0usize;
        while bytes.len().saturating_sub(offset) >= 8 {
            let header = &bytes[offset..offset + 8];
            let container = header[0] & 0xF == 0xF;
            let kind = u16::from_le_bytes([header[2], header[3]]);
            let len = u32::from_le_bytes([header[4], header[5], header[6], header[7]]) as usize;
            let payload = offset
                .checked_add(8)
                .and_then(|start| bytes.get(start..start.checked_add(len)?))?;
            let found = match kind {
                // `refType` 1 is a shape (`ElementTypeEnum` ET_ShapeType).
                VISUAL_SHAPE_ATOM
                    if at(payload, 4) == Some(1) && at(payload, 8) == Some(shape_id) =>
                {
                    Some("an animation of the slide targets the equation's frame")
                },
                BUILD_ATOM if at(payload, 8) == Some(shape_id) => {
                    Some("a build list of the slide names the equation's frame")
                },
                CONNECTOR_RULE
                    if [4, 8, 12]
                        .iter()
                        .any(|field| at(payload, *field) == Some(shape_id)) =>
                {
                    Some("a connector of the slide is attached to the equation's frame")
                },
                ARC_RULE | CALLOUT_RULE if at(payload, 4) == Some(shape_id) => {
                    Some("a drawing rule of the slide names the equation's frame")
                },
                _ if (container || kind == BINARY_TAG_DATA_BLOB) && tree::is_sequence(payload) => {
                    scan(payload, shape_id, depth + 1)
                },
                _ => None,
            };
            if found.is_some() {
                return found;
            }
            offset += 8 + len;
        }
        None
    }
    scan(record, shape_id, 0)
}

/// The patriarch group of an expanded slide tree.
#[cfg(test)]
pub(super) fn patriarch_for_tests(slide: &mut Node) -> &mut Node {
    patriarch_mut(slide).expect("the slide has a drawing")
}

/// The child-index path from `group` to the shape with `shape_id`.
fn find_shape(group: &Node, shape_id: u32, depth: usize) -> Option<Vec<usize>> {
    if depth > 32 {
        return None;
    }
    for (index, child) in group.children().iter().enumerate() {
        match child.record_type() {
            SP_CONTAINER if read::shape_id(child) == Some(shape_id) => return Some(vec![index]),
            SPGR_CONTAINER => {
                if let Some(mut path) = find_shape(child, shape_id, depth + 1) {
                    path.insert(0, index);
                    return Some(path);
                }
            },
            _ => {},
        }
    }
    None
}

/// Every shape identifier of a drawing.
pub(super) fn shape_ids(node: &Node, out: &mut Vec<u32>, depth: usize) {
    if depth > 64 {
        return;
    }
    for child in node.children() {
        if child.record_type() == FSP
            && let Some(payload) = child.payload()
            && payload.len() >= 4
        {
            out.push(u32::from_le_bytes([
                payload[0], payload[1], payload[2], payload[3],
            ]));
        }
        shape_ids(child, out, depth + 1);
    }
}

/// Add or remove one shape in a slide drawing's `OfficeArtFDG` count.
fn adjust_shape_count(slide: &mut Node, add: bool) -> Result<()> {
    let Some(fdg) = slide
        .child_mut(PP_DRAWING)?
        .and_then(|drawing| drawing.child_mut(DG_CONTAINER).ok().flatten())
        .and_then(|container| container.child_mut(FDG).ok().flatten())
        .and_then(Node::payload_mut)
        .filter(|payload| payload.len() >= 8)
    else {
        return Ok(());
    };
    let count = u32::from_le_bytes([fdg[0], fdg[1], fdg[2], fdg[3]]);
    let count = if add {
        count.saturating_add(1)
    } else {
        count.saturating_sub(1)
    };
    fdg[..4].copy_from_slice(&count.to_le_bytes());
    Ok(())
}

/// The preview of `native` at `options`, drawn at a smaller scale when the
/// picture would exceed the pixel ceiling: a long equation gets a coarser
/// picture of the same size in points rather than none.
fn render_fitting(
    native: &[u8],
    mut options: PngOptions,
    limits: &Limits,
) -> std::result::Result<litchi_equation::Png, litchi_equation::Error> {
    loop {
        match litchi_equation::mtef_to_png(native, &options, limits) {
            Err(litchi_equation::Error::Render(_)) if options.scale() > MIN_PREVIEW_SCALE => {
                options = options.with_scale((options.scale() / 2.0).max(MIN_PREVIEW_SCALE))?;
            },
            result => return result,
        }
    }
}

/// A frame at (`left`, `top`) of `width` × `height` points.
fn frame_bounds(left: i32, top: i32, width: f64, height: f64) -> Option<Bounds> {
    let extent = |points: f64| -> Option<i32> {
        let units = (points * MASTER_UNITS_PER_POINT).round();
        (units.is_finite() && units >= 1.0 && units <= f64::from(i16::MAX)).then_some(units as i32)
    };
    let right = left.checked_add(extent(width)?)?;
    let bottom = top.checked_add(extent(height)?)?;
    let fits = |value: i32| i16::try_from(value).is_ok();
    (fits(left) && fits(top) && fits(right) && fits(bottom))
        .then(|| Bounds::from_edges(left, top, right, bottom))
}

/// `bounds` with its top-left corner kept and its size scaled by the size
/// change from `old` to `new` (points); when the old equation cannot be
/// measured, the height is kept and the width follows the new aspect
/// ratio.
fn rescaled(
    bounds: Bounds,
    old: Option<(f64, f64)>,
    new: (f64, f64),
) -> std::result::Result<Bounds, ()> {
    let width = bounds.width() as f64;
    let height = bounds.height() as f64;
    let (width, height) = match old {
        Some((old_width, old_height)) if old_width > 0.0 && old_height > 0.0 => {
            (width * new.0 / old_width, height * new.1 / old_height)
        },
        _ if new.1 > 0.0 => (height * new.0 / new.1, height),
        _ => (width, height),
    };
    let extent = |value: f64| -> std::result::Result<i32, ()> {
        let value = value.round();
        if value.is_finite() && value >= 1.0 && value <= f64::from(i32::MAX) {
            Ok(value as i32)
        } else {
            Err(())
        }
    };
    let right = bounds.left.checked_add(extent(width)?).ok_or(())?;
    let bottom = bounds.top.checked_add(extent(height)?).ok_or(())?;
    Ok(Bounds::from_edges(bounds.left, bounds.top, right, bottom))
}

/// The `Equation Native` stream of `source`, the object kind it is written
/// as, and the conversion losses. `kind` is the existing object's MTEF when
/// replacing one.
fn equation_native(
    source: Source<'_>,
    kind: Option<MtefTarget>,
    limits: &Limits,
) -> Result<(Vec<u8>, MtefTarget, Vec<String>)> {
    let expected = kind.as_ref().map(version_of);
    let target = kind.unwrap_or(MtefTarget::MathType);
    match source {
        Source::Latex(latex) => {
            let native = litchi_equation::latex_to_equation_native(latex, &target, limits)?;
            let losses = native.losses().to_vec();
            Ok((native.into_bytes(), target, losses))
        },
        Source::Omml(omml) => {
            let native = litchi_equation::omml_to_equation_native(omml, &target, limits)?;
            let losses = native.losses().to_vec();
            Ok((native.into_bytes(), target, losses))
        },
        Source::Mtef(bytes) => {
            let read = litchi_equation::read_mtef(bytes, limits)?;
            let info = read.info().cloned().ok_or_else(|| {
                refused(
                    RefusalKind::InvalidValue,
                    "equation.mtef",
                    "the MTEF has no header",
                )
            })?;
            if expected
                .is_some_and(|expected| expected != version_of(&MtefTarget::SameAs(info.clone())))
            {
                return Err(refused(
                    RefusalKind::InvalidValue,
                    "equation.mtef",
                    "the MTEF version differs from the equation's",
                ));
            }
            let native = if info.ole_header.is_some() {
                bytes.to_vec()
            } else {
                with_ole_header(bytes)?
            };
            Ok((native, MtefTarget::SameAs(info), Vec::new()))
        },
    }
}

/// The MTEF version an object of `target`'s kind holds: 3 or 5.
fn version_of(target: &MtefTarget) -> u8 {
    if target.prog_id() == "Equation.3" {
        3
    } else {
        5
    }
}

/// Bare MTEF behind an `EQNOLEFILEHDR` ([MathType SDK] "Equation Native"
/// header: `cbHdr` 28, version 0x00020000, the MathType EF clipboard
/// format, the MTEF length and four reserved zero fields).
fn with_ole_header(mtef: &[u8]) -> Result<Vec<u8>> {
    let len = u32::try_from(mtef.len())
        .map_err(|_error| PackageError::ResourceLimit("the MTEF exceeds 4 GiB".into()))?;
    let mut out = Vec::with_capacity(OLE_HEADER_LEN + mtef.len());
    out.extend_from_slice(&28u16.to_le_bytes());
    out.extend_from_slice(&0x0002_0000u32.to_le_bytes());
    out.extend_from_slice(&0xC1C6u16.to_le_bytes());
    out.extend_from_slice(&len.to_le_bytes());
    out.extend_from_slice(&[0; 16]);
    out.extend_from_slice(mtef);
    Ok(out)
}

impl Package {
    /// Start an edit of the equations on this presentation's slides.
    ///
    /// # Errors
    ///
    /// An encrypted presentation is refused; a presentation whose slide
    /// list, object list or persisted records cannot be read reports the
    /// read error.
    pub fn edit_equations(&self) -> Result<Edit> {
        Edit::new(self)
    }

    /// Apply an equation patch to the exact package it was made from.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// source.
    pub fn apply_equations(&self, patch: &Patch) -> Result<Commit> {
        super::apply_patch(self, patch)
    }
}
