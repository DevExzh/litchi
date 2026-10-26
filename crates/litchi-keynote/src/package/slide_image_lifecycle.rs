//! Exact-source, selector-first lifecycle of ordinary slide images (record
//! 0372; a port of the migration host's `duplicate_slide_image`,
//! `flip_slide_image`, `restore_slide_image_original_size`,
//! `set_slide_image_properties`, `replace_slide_image_data` and the image
//! reads).
//!
//! An image is addressed by its slide and an [`ImageSelector`] (its
//! position among the slide's images in source order). An [`Edit`] stages
//! operations against one immutable package; [`Edit::commit`] applies them
//! in one structural workspace, publishes one candidate, reopens it and
//! reads every touched image back. Images materialized from the slide's
//! layout are read-only here.

use std::collections::{BTreeSet, HashMap};
use std::fmt;
use std::sync::Arc;

use litchi_iwa_archive::package::ExactArtifacts;
use litchi_iwa_common::WireLimits;
use litchi_iwa_common::media::Type as MediaType;
use litchi_iwa_common::shape::geometry::{FlipAxis, Size};
use sha1::{Digest, Sha1};
use thiserror::Error as ThisError;

use super::image_adjustments::{self, SelectedImage, select_slide_image};
use super::slide_drawable::Geometry;
use super::structure::StructureError;
use super::structure::wire::Fields;
use super::structure::workspace::Workspace;
use super::{Package, PhysicalSource, ReadError};
use crate::slide::media::MediaProperties;
use crate::{ImageSelector, SlideSelector};

const IMAGE_MESSAGE_TYPE: u32 = 3_005;
/// `TSD.ImageArchive.super` (`TSD.DrawableArchive`).
const IMAGE_DRAWABLE: u32 = 1;
const IMAGE_ORIGINAL_SIZE: u32 = 4;
const IMAGE_NATURAL_SIZE: u32 = 9;
const IMAGE_DATA: u32 = 11;
const DRAWABLE_GEOMETRY: u32 = 1;
const DRAWABLE_HYPERLINK_URL: u32 = 4;
const DRAWABLE_LOCKED: u32 = 5;
const DRAWABLE_ASPECT_RATIO_LOCKED: u32 = 7;
const DRAWABLE_ACCESSIBILITY_DESCRIPTION: u32 = 8;
const GEOMETRY_FLAGS: u32 = 3;
const GEOMETRY_ANGLE: u32 = 4;
/// The `TSD.GeometryArchive.flags` bit Keynote's Arrange Flip toggles.
const REFLECTION_FLAG: u64 = 1 << 2;
const DUPLICATE_OFFSET: f32 = 10.0;
const HALF_TURN_DEGREES: f32 = 180.0;
const FULL_TURN_DEGREES: f32 = 360.0;

/// A finite resource governed by an image-lifecycle transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum LimitKind {
    /// Native objects in the package.
    Objects,
    /// Bytes of one replacement image.
    DataBytes,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(match self {
            Self::Objects => "objects",
            Self::DataBytes => "data bytes",
        })
    }
}

/// An image-lifecycle failure that never exposes native identifiers.
#[derive(Debug, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses.
    #[error(transparent)]
    Refused(Refusal),
    /// The slide or image selector did not resolve.
    #[error(transparent)]
    Selection(image_adjustments::Error),
    /// The package cannot be edited safely.
    #[error("the Keynote source cannot be edited safely")]
    InvalidSource,
    /// A resource limit was exceeded.
    #[error(
        "Keynote image-lifecycle {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    Limit {
        /// The exceeded resource.
        kind: LimitKind,
        /// The observed amount.
        observed: u64,
        /// The configured maximum.
        maximum: u64,
    },
    /// The published candidate did not read back as staged.
    #[error("the Keynote image-lifecycle candidate failed verification")]
    Verification,
    /// The patch does not match the exact source package.
    #[error("the Keynote image-lifecycle patch does not match the exact source package")]
    PatchConflict,
}

/// A well-formed request the image-lifecycle family refuses.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The package was not opened from an exact physical source.
    #[error("this Keynote source does not support structural edits")]
    UnsupportedSource,
    /// The image was materialized from the slide's layout.
    #[error("the Keynote image belongs to the slide layout")]
    LayoutImage,
    /// The image's frame records no native flags, so it cannot be flipped.
    #[error("the Keynote image frame records no flip state")]
    UnflippableGeometry,
    /// The image records no usable original size.
    #[error("the Keynote image records no original size")]
    NoOriginalSize,
    /// The replacement image is empty.
    #[error("the replacement image is empty")]
    EmptyData,
    /// The replacement's detected media type differs from the image's.
    #[error("the replacement is not the same kind of media as the image")]
    DataType,
    /// The image's graph is referenced from outside the image.
    #[error("the Keynote image is referenced from outside its private graph")]
    SharedGraph,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::UnsupportedSource
            | Self::LayoutImage
            | Self::UnflippableGeometry
            | Self::NoOriginalSize
            | Self::SharedGraph => RefusalKind::UnsupportedContent,
            Self::EmptyData | Self::DataType => RefusalKind::InvalidValue,
        }
    }
}

/// Whether a slide image is an ordinary file image or materialized from
/// the slide's layout.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Kind {
    /// An ordinary, independently editable image.
    File,
    /// An image materialized from the slide's layout (read-only).
    Layout,
}

/// A read-only, archive-free view of one slide image.
#[derive(Debug, Clone, PartialEq)]
pub struct Info {
    kind: Kind,
    geometry: Option<Geometry>,
    reflected: bool,
    properties: MediaProperties,
    original_size: Option<Size>,
    natural_size: Option<Size>,
}

impl Info {
    /// Whether the image is a file or layout image.
    #[must_use]
    pub const fn kind(&self) -> Kind {
        self.kind
    }

    /// The displayed frame, when recorded.
    #[must_use]
    pub const fn geometry(&self) -> Option<Geometry> {
        self.geometry
    }

    /// Whether the image is reflected by an Arrange Flip.
    #[must_use]
    pub const fn is_reflected(&self) -> bool {
        self.reflected
    }

    /// Hyperlink, lock and accessibility properties.
    #[must_use]
    pub const fn properties(&self) -> &MediaProperties {
        &self.properties
    }

    /// The stored original media size, the target of
    /// [`Edit::restore_original_size`].
    #[must_use]
    pub const fn original_size(&self) -> Option<Size> {
        self.original_size
    }

    /// The untransformed media size.
    #[must_use]
    pub const fn natural_size(&self) -> Option<Size> {
        self.natural_size
    }
}

#[derive(Clone)]
enum Change {
    Duplicate,
    Flip(FlipAxis),
    RestoreOriginalSize,
    Properties(MediaProperties),
    Data(Arc<[u8]>),
}

#[derive(Clone)]
struct Operation {
    selected: SelectedImage,
    change: Change,
}

/// One bounded image-lifecycle edit staged against an immutable package.
///
/// ```no_run
/// use litchi_keynote::slide::image::lifecycle::FlipAxis;
/// use litchi_keynote::{ImageSelector, Package, SlideSelector};
///
/// # fn main() -> Result<(), Box<dyn std::error::Error>> {
/// let package = Package::open("deck.key")?;
/// let mut edit = package.edit_slide_image_lifecycle();
/// edit.flip(SlideSelector::index(0), ImageSelector::index(0), FlipAxis::Horizontal)?;
/// let _commit = edit.commit()?;
/// # Ok(())
/// # }
/// ```
pub struct Edit<'a> {
    source: &'a Package,
    operations: Vec<Operation>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Edit")
            .field("operations", &self.operations.len())
            .finish_non_exhaustive()
    }
}

impl<'a> Edit<'a> {
    const fn new(source: &'a Package) -> Self {
        Self {
            source,
            operations: Vec::new(),
        }
    }

    fn stage(
        &mut self,
        slide: SlideSelector<'_>,
        image: ImageSelector,
        change: Change,
    ) -> Result<&mut Self, Error> {
        let selected = select_slide_image(self.source, slide, image).map_err(Error::Selection)?;
        if selected.layout {
            return Err(Error::Refused(Refusal::LayoutImage));
        }
        self.operations.push(Operation { selected, change });
        Ok(self)
    }

    /// Stage a copy of one image on its slide, offset by ten points (host
    /// `duplicate_slide_image`). The copy and its private title and caption
    /// get fresh identities; the embedded image data stays shared, as
    /// Keynote's Duplicate shares it.
    ///
    /// # Errors
    ///
    /// Returns typed selector errors and [`Refusal::LayoutImage`].
    pub fn duplicate<'s>(
        &mut self,
        slide: impl Into<SlideSelector<'s>>,
        image: impl Into<ImageSelector>,
    ) -> Result<&mut Self, Error> {
        self.stage(slide.into(), image.into(), Change::Duplicate)
    }

    /// Stage one native Arrange Flip (host `flip_slide_image`): the
    /// reflection toggles; a vertical flip also adds a half turn.
    ///
    /// # Errors
    ///
    /// Returns typed selector errors and [`Refusal::LayoutImage`];
    /// [`Edit::commit`] refuses a frame without native flags.
    pub fn flip<'s>(
        &mut self,
        slide: impl Into<SlideSelector<'s>>,
        image: impl Into<ImageSelector>,
        axis: FlipAxis,
    ) -> Result<&mut Self, Error> {
        self.stage(slide.into(), image.into(), Change::Flip(axis))
    }

    /// Stage restoring the displayed size from the stored original size,
    /// keeping position, rotation and reflection (host
    /// `restore_slide_image_original_size`).
    ///
    /// # Errors
    ///
    /// Returns typed selector errors and [`Refusal::LayoutImage`];
    /// [`Edit::commit`] refuses an image without an original size.
    pub fn restore_original_size<'s>(
        &mut self,
        slide: impl Into<SlideSelector<'s>>,
        image: impl Into<ImageSelector>,
    ) -> Result<&mut Self, Error> {
        self.stage(slide.into(), image.into(), Change::RestoreOriginalSize)
    }

    /// Stage the image's hyperlink, lock and accessibility properties (host
    /// `set_slide_image_properties`); `None` omits a property.
    ///
    /// # Errors
    ///
    /// Returns typed selector errors and [`Refusal::LayoutImage`].
    pub fn set_properties<'s>(
        &mut self,
        slide: impl Into<SlideSelector<'s>>,
        image: impl Into<ImageSelector>,
        properties: MediaProperties,
    ) -> Result<&mut Self, Error> {
        self.stage(slide.into(), image.into(), Change::Properties(properties))
    }

    /// Stage replacing the image's embedded data (host
    /// `replace_slide_image_data`). Images sharing the data observe the
    /// replacement, as in Keynote.
    ///
    /// # Errors
    ///
    /// Returns typed selector errors, [`Refusal::LayoutImage`] and
    /// [`Refusal::EmptyData`]; [`Edit::commit`] refuses a replacement of a
    /// different media type and one beyond the package's entry limit.
    pub fn set_data<'s>(
        &mut self,
        slide: impl Into<SlideSelector<'s>>,
        image: impl Into<ImageSelector>,
        data: &[u8],
    ) -> Result<&mut Self, Error> {
        if data.is_empty() {
            return Err(Error::Refused(Refusal::EmptyData));
        }
        self.stage(slide.into(), image.into(), Change::Data(Arc::from(data)))
    }

    /// Apply every staged operation, publish one candidate, reopen it and
    /// read the touched images back.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Verification`] when the candidate does not read back
    /// as staged and typed source errors.
    pub fn commit(self) -> Result<Commit, Error> {
        commit_operations(self.source, &self.operations)
    }
}

fn lift(_error: StructureError) -> Error {
    Error::InvalidSource
}

fn physical_source(package: &Package) -> Result<Arc<[u8]>, Error> {
    match &package.state.source {
        PhysicalSource::Package(source) if source.source_is_exact() => Ok(source.shared_source()),
        PhysicalSource::Package(_) | PhysicalSource::Semantic(_) => {
            Err(Error::Refused(Refusal::UnsupportedSource))
        },
    }
}

fn map_read_error(error: ReadError) -> Error {
    match error {
        ReadError::SemanticLimit {
            observed, maximum, ..
        } => Error::Limit {
            kind: LimitKind::Objects,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        _ => Error::InvalidSource,
    }
}

/// The modelled values of one `TSD.ImageArchive` payload.
struct Values {
    geometry: Option<Geometry>,
    flags: Option<u64>,
    properties: MediaProperties,
    original_size: Option<Size>,
    natural_size: Option<Size>,
    data: Option<u64>,
}

fn size_of(fields: &Fields, number: u32, wire: WireLimits) -> Result<Option<Size>, Error> {
    let Some(field) = fields.unique(number).map_err(lift)? else {
        return Ok(None);
    };
    let inner = Fields::parse(field.message().map_err(lift)?, wire).map_err(lift)?;
    let value = |number| -> Result<Option<f32>, Error> {
        inner
            .unique(number)
            .map_err(lift)?
            .map(|field| field.float().map_err(lift))
            .transpose()
    };
    Ok(value(1)?
        .zip(value(2)?)
        .map(|(width, height)| Size { width, height }))
}

fn image_values(payload: &[u8], wire: WireLimits) -> Result<Values, Error> {
    let fields = Fields::parse(payload, wire).map_err(lift)?;
    let drawable = fields
        .unique(IMAGE_DRAWABLE)
        .map_err(lift)?
        .ok_or(Error::InvalidSource)?
        .message()
        .map_err(lift)?;
    let geometry = super::slide_drawable::read_geometry(drawable, wire).map_err(lift)?;
    let drawable = Fields::parse(drawable, wire).map_err(lift)?;
    let flags = match drawable.unique(DRAWABLE_GEOMETRY).map_err(lift)? {
        Some(geometry) => Fields::parse(geometry.message().map_err(lift)?, wire)
            .map_err(lift)?
            .unique(GEOMETRY_FLAGS)
            .map_err(lift)?
            .map(|field| field.varint().map_err(lift))
            .transpose()?,
        None => None,
    };
    let text = |number| -> Result<Option<String>, Error> {
        drawable
            .unique(number)
            .map_err(lift)?
            .map(|field| {
                std::str::from_utf8(field.message().map_err(lift)?)
                    .map(str::to_owned)
                    .map_err(|_| Error::InvalidSource)
            })
            .transpose()
    };
    let flag = |number| -> Result<Option<bool>, Error> {
        drawable
            .unique(number)
            .map_err(lift)?
            .map(|field| field.varint().map(|value| value != 0).map_err(lift))
            .transpose()
    };
    let properties = MediaProperties::from_parts(
        text(DRAWABLE_HYPERLINK_URL)?,
        flag(DRAWABLE_LOCKED)?,
        flag(DRAWABLE_ASPECT_RATIO_LOCKED)?,
        text(DRAWABLE_ACCESSIBILITY_DESCRIPTION)?,
    );
    let data = match fields.unique(IMAGE_DATA).map_err(lift)? {
        Some(reference) => Fields::parse(reference.message().map_err(lift)?, wire)
            .map_err(lift)?
            .unique(1)
            .map_err(lift)?
            .map(|field| field.varint().map_err(lift))
            .transpose()?,
        None => None,
    };
    Ok(Values {
        geometry,
        flags,
        properties,
        original_size: size_of(&fields, IMAGE_ORIGINAL_SIZE, wire)?,
        natural_size: size_of(&fields, IMAGE_NATURAL_SIZE, wire)?,
        data,
    })
}

fn object_payload(package: &Package, identifier: u64) -> Result<&[u8], Error> {
    package
        .object(identifier)
        .ok_or(Error::InvalidSource)?
        .messages
        .iter()
        .find(|message| message.type_ == IMAGE_MESSAGE_TYPE)
        .map(|message| message.data.as_slice())
        .ok_or(Error::InvalidSource)
}

fn info_of(package: &Package, selected: SelectedImage) -> Result<Info, Error> {
    let wire = package.wire_limits().map_err(|_| Error::InvalidSource)?;
    let values = image_values(object_payload(package, selected.image_identifier)?, wire)?;
    Ok(Info {
        kind: if selected.layout {
            Kind::Layout
        } else {
            Kind::File
        },
        geometry: values.geometry,
        reflected: values
            .flags
            .is_some_and(|flags| flags & REFLECTION_FLAG != 0),
        properties: values.properties,
        original_size: values.original_size,
        natural_size: values.natural_size,
    })
}

/// The package member holding data item `data`.
fn data_member(workspace: &Workspace<'_>, data: u64) -> Result<String, Error> {
    let name = workspace
        .registry
        .data_file_name(data, workspace.wire)
        .map_err(lift)?
        .ok_or(Error::InvalidSource)?;
    Ok(format!("Data/{name}"))
}

/// Edit the image payload's nested message at `path` and store it.
fn edit_image(
    workspace: &mut Workspace<'_>,
    image: u64,
    path: &[u32],
    edit: &mut dyn FnMut(&mut Fields) -> Result<(), StructureError>,
) -> Result<(), Error> {
    let wire = workspace.wire;
    let payload = workspace
        .payload(image, IMAGE_MESSAGE_TYPE)
        .map_err(lift)?
        .to_vec();
    let data =
        super::slide_drawable::edit_nested_payload(&payload, path, wire, edit).map_err(lift)?;
    super::slide_drawable::replace_message_payload(workspace, image, IMAGE_MESSAGE_TYPE, data)
        .map_err(lift)
}

fn set_text_field(fields: &mut Fields, number: u32, value: Option<&str>) {
    match value {
        Some(value) => fields.set_message(number, value.as_bytes()),
        None => {
            fields.remove(number);
        },
    }
}

fn set_flag_field(fields: &mut Fields, number: u32, value: Option<bool>) {
    match value {
        Some(value) => fields.set_varint(number, u64::from(value)),
        None => {
            fields.remove(number);
        },
    }
}

fn vertical_flip_angle(angle: Option<f32>) -> f32 {
    let current = angle.unwrap_or_default().rem_euclid(FULL_TURN_DEGREES);
    let flipped = (current + HALF_TURN_DEGREES).rem_euclid(FULL_TURN_DEGREES);
    if flipped == 0.0 { 0.0 } else { flipped }
}

fn apply(
    workspace: &mut Workspace<'_>,
    source: &Package,
    operation: &Operation,
) -> Result<(), Error> {
    let image = operation.selected.image_identifier;
    let wire = workspace.wire;
    let values = image_values(
        workspace.payload(image, IMAGE_MESSAGE_TYPE).map_err(lift)?,
        wire,
    )?;
    match &operation.change {
        Change::Duplicate => duplicate_image(workspace, operation.selected.slide_identifier, image),
        &Change::Flip(axis) => {
            let flags = values
                .flags
                .filter(|_| values.geometry.is_some_and(Geometry::is_valid))
                .ok_or(Error::Refused(Refusal::UnflippableGeometry))?;
            let angle = values.geometry.and_then(|geometry| geometry.rotation);
            let angle = match axis {
                FlipAxis::Horizontal => angle,
                FlipAxis::Vertical => Some(vertical_flip_angle(angle)),
            };
            edit_image(
                workspace,
                image,
                &[IMAGE_DRAWABLE, DRAWABLE_GEOMETRY],
                &mut |fields| {
                    fields.set_varint(GEOMETRY_FLAGS, flags ^ REFLECTION_FLAG);
                    if let Some(angle) = angle {
                        fields.set_float(GEOMETRY_ANGLE, angle);
                    }
                    Ok(())
                },
            )
        },
        Change::RestoreOriginalSize => {
            let size = values
                .original_size
                .filter(|size| {
                    size.width.is_finite()
                        && size.height.is_finite()
                        && size.width > 0.0
                        && size.height > 0.0
                })
                .ok_or(Error::Refused(Refusal::NoOriginalSize))?;
            let geometry = values
                .geometry
                .filter(|geometry| geometry.is_valid())
                .ok_or(Error::InvalidSource)?;
            let restored = Geometry {
                width: size.width,
                height: size.height,
                ..geometry
            };
            edit_image(
                workspace,
                image,
                &[IMAGE_DRAWABLE, DRAWABLE_GEOMETRY],
                &mut |fields| super::slide_drawable::patch_geometry_fields(fields, restored, wire),
            )
        },
        Change::Properties(properties) => {
            edit_image(workspace, image, &[IMAGE_DRAWABLE], &mut |fields| {
                set_text_field(fields, DRAWABLE_HYPERLINK_URL, properties.hyperlink_url());
                set_flag_field(fields, DRAWABLE_LOCKED, properties.locked());
                set_flag_field(
                    fields,
                    DRAWABLE_ASPECT_RATIO_LOCKED,
                    properties.aspect_ratio_locked(),
                );
                set_text_field(
                    fields,
                    DRAWABLE_ACCESSIBILITY_DESCRIPTION,
                    properties.accessibility_description(),
                );
                Ok(())
            })
        },
        Change::Data(replacement) => {
            let data = values.data.ok_or(Error::InvalidSource)?;
            let member = data_member(workspace, data)?;
            let before = source
                .structure_member_bytes(&member)
                .ok_or(Error::InvalidSource)?;
            let maximum = source.state.options.archive().max_entry_bytes();
            let observed = replacement.len() as u64;
            if observed > maximum {
                return Err(Error::Limit {
                    kind: LimitKind::DataBytes,
                    observed,
                    maximum,
                });
            }
            let (before_type, after_type) = (
                MediaType::from_bytes(before),
                MediaType::from_bytes(replacement),
            );
            if before_type != MediaType::Unknown
                && after_type != MediaType::Unknown
                && before_type != after_type
            {
                return Err(Error::Refused(Refusal::DataType));
            }
            let digest = Sha1::digest(replacement.as_ref());
            workspace
                .registry
                .patch_data_info(data, digest.as_slice(), observed, wire)
                .map_err(lift)?
                .ok_or(Error::InvalidSource)?;
            workspace
                .replace_member(&member, replacement.to_vec())
                .map_err(lift)
        },
    }
}

/// The image's private graph inside its component, in archive order (the
/// host's `private_clone_object_ids` order).
fn image_graph(workspace: &Workspace<'_>, slide: u64, image: u64) -> Result<Vec<u64>, Error> {
    let member = workspace.component_of(image).ok_or(Error::InvalidSource)?;
    let order = workspace.member_objects(&member);
    let internal = order.iter().copied().collect::<BTreeSet<_>>();
    let mut closure = BTreeSet::from([image]);
    let mut pending = vec![image];
    while let Some(current) = pending.pop() {
        for target in
            super::slide_drawable::outgoing_references(workspace, current).map_err(lift)?
        {
            if target == slide {
                // The parent edge lives in the payload; the header graph
                // the host clones never reaches the slide.
                continue;
            }
            if internal.contains(&target) && closure.insert(target) {
                pending.push(target);
            }
        }
    }
    Ok(order
        .into_iter()
        .filter(|object| closure.contains(object))
        .collect())
}

/// Duplicate one image (host `duplicate_slide_image`).
fn duplicate_image(workspace: &mut Workspace<'_>, slide: u64, image: u64) -> Result<(), Error> {
    let wire = workspace.wire;
    let member = workspace.component_of(image).ok_or(Error::InvalidSource)?;
    let objects = image_graph(workspace, slide, image)?;
    let mut remap = HashMap::with_capacity(objects.len());
    for object in &objects {
        remap.insert(*object, workspace.allocate().map_err(lift)?);
    }
    for object in &objects {
        let original = workspace.require(*object).map_err(lift)?.clone();
        let clone = workspace
            .clone_object(&original, remap[object], &remap)
            .map_err(lift)?;
        workspace.insert(&member, clone).map_err(lift)?;
    }
    let copy = remap[&image];
    let geometry = image_values(
        workspace.payload(copy, IMAGE_MESSAGE_TYPE).map_err(lift)?,
        wire,
    )?
    .geometry
    .ok_or(Error::InvalidSource)?;
    let offset = Geometry {
        x: geometry.x + DUPLICATE_OFFSET,
        y: geometry.y + DUPLICATE_OFFSET,
        ..geometry
    };
    edit_image(
        workspace,
        copy,
        &[IMAGE_DRAWABLE, DRAWABLE_GEOMETRY],
        &mut |fields| super::slide_drawable::patch_geometry_fields(fields, offset, wire),
    )?;
    super::slide_drawable::attach_drawable(workspace, slide, copy).map_err(lift)?;
    let owner = workspace
        .registry
        .component_for_member(&member)
        .ok_or(Error::InvalidSource)?
        .clone();
    let mut fresh = Vec::new();
    for object in &objects {
        if let Some(uuid) = owner.uuid_of(*object) {
            fresh.push((
                remap[object],
                workspace
                    .clone_uuid(Some(uuid), remap[object])
                    .map_err(lift)?,
            ));
        }
    }
    let mut owners = Vec::new();
    for object in &objects {
        for datum in owner.data_owned_by(&[*object]) {
            owners.push((datum, remap[object]));
        }
    }
    let entry = workspace
        .registry
        .component_mut(owner.identifier)
        .ok_or(Error::InvalidSource)?;
    for (object, uuid) in fresh {
        entry.add_object_uuid(object, uuid);
    }
    for (datum, object) in owners {
        entry.add_data_owner(datum, object);
    }
    Ok(())
}

fn verify(candidate: &Package, operation: &Operation, before: &Info) -> Result<(), Error> {
    let selected = select_slide_image(
        candidate,
        SlideSelector::position(operation.selected.slide_position),
        ImageSelector::position(operation.selected.image_position),
    )
    .map_err(|_error| Error::Verification)?;
    let after = info_of(candidate, selected)?;
    let ok = match &operation.change {
        Change::Duplicate => after == *before,
        Change::Flip(_) => after.reflected != before.reflected,
        Change::RestoreOriginalSize => {
            after
                .geometry
                .zip(after.original_size)
                .is_some_and(|(geometry, size)| {
                    (geometry.width, geometry.height) == (size.width, size.height)
                })
        },
        Change::Properties(properties) => after.properties == *properties,
        Change::Data(replacement) => {
            let workspace = Workspace::new(candidate).map_err(lift)?;
            let wire = workspace.wire;
            let data = image_values(object_payload(candidate, selected.image_identifier)?, wire)?
                .data
                .ok_or(Error::Verification)?;
            candidate.structure_member_bytes(&data_member(&workspace, data)?)
                == Some(replacement.as_ref())
        },
    };
    if ok { Ok(()) } else { Err(Error::Verification) }
}

fn commit_operations(source: &Package, operations: &[Operation]) -> Result<Commit, Error> {
    let source_bytes = physical_source(source)?;
    if operations.is_empty() {
        return Ok(Commit {
            package: source.clone(),
            patch: Patch {
                artifacts: ExactArtifacts::new(Arc::clone(&source_bytes), source_bytes),
            },
            diagnostics: Diagnostics::default(),
        });
    }
    let before = operations
        .iter()
        .map(|operation| info_of(source, operation.selected))
        .collect::<Result<Vec<_>, _>>()?;
    let mut workspace = Workspace::new(source).map_err(lift)?;
    for operation in operations {
        apply(&mut workspace, source, operation)?;
    }
    let bytes = workspace.publish().map_err(lift)?;
    let candidate = source.structure_reopen(bytes).map_err(map_read_error)?;
    for (operation, before) in operations.iter().zip(&before) {
        verify(&candidate, operation, before)?;
    }
    let target = physical_source(&candidate)?;
    let patch = Patch {
        artifacts: ExactArtifacts::new(source_bytes, target),
    };
    Ok(Commit {
        package: candidate,
        diagnostics: Diagnostics {
            changed: !patch.is_noop(),
            operations: operations.len(),
        },
        patch,
    })
}

/// Exact-source-checked reversible image-lifecycle patch.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    artifacts: ExactArtifacts,
}

impl fmt::Debug for Patch {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Patch").finish_non_exhaustive()
    }
}

impl Patch {
    /// Whether the patch leaves the package unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.artifacts.is_byte_noop()
    }

    /// The patch that restores the source package.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
        }
    }
}

/// What an image-lifecycle commit changed.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
    operations: usize,
}

impl Diagnostics {
    /// Whether the package changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// How many staged operations were applied.
    #[must_use]
    pub const fn operations(self) -> usize {
        self.operations
    }
}

/// A published image-lifecycle commit.
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// The edited package.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Take the edited package.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// The reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// What the commit changed.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }

    /// Split into the edited package and the patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }
}

impl Package {
    /// Begin an image-lifecycle edit of this package.
    #[must_use]
    pub const fn edit_slide_image_lifecycle(&self) -> Edit<'_> {
        Edit::new(self)
    }

    /// Read one slide image (host `slide_images`, `slide_image_properties`).
    ///
    /// # Errors
    ///
    /// Returns typed selector errors and [`Error::InvalidSource`].
    pub fn slide_image<'s>(
        &self,
        slide: impl Into<SlideSelector<'s>>,
        image: impl Into<ImageSelector>,
    ) -> Result<Info, Error> {
        let selected =
            select_slide_image(self, slide.into(), image.into()).map_err(Error::Selection)?;
        info_of(self, selected)
    }

    /// Every image of one slide, in source order (host `slide_images`).
    ///
    /// # Errors
    ///
    /// Returns typed selector errors and [`Error::InvalidSource`].
    pub fn slide_image_list<'s>(
        &self,
        slide: impl Into<SlideSelector<'s>>,
    ) -> Result<Vec<Info>, Error> {
        let slide = slide.into();
        let mut images = Vec::new();
        loop {
            match select_slide_image(self, slide, ImageSelector::index(images.len())) {
                Ok(selected) => images.push(info_of(self, selected)?),
                Err(image_adjustments::Error::Refused(
                    image_adjustments::Refusal::ImagePositionNotFound { .. },
                )) => return Ok(images),
                Err(error) => return Err(Error::Selection(error)),
            }
        }
    }

    /// The embedded data of one slide image (host `extract_media` of the
    /// image's data).
    ///
    /// # Errors
    ///
    /// Returns typed selector errors and [`Error::InvalidSource`] when the
    /// data is not materialized in the package.
    pub fn slide_image_data<'s>(
        &self,
        slide: impl Into<SlideSelector<'s>>,
        image: impl Into<ImageSelector>,
    ) -> Result<&[u8], Error> {
        let selected =
            select_slide_image(self, slide.into(), image.into()).map_err(Error::Selection)?;
        let workspace = Workspace::new(self).map_err(lift)?;
        let data = image_values(
            object_payload(self, selected.image_identifier)?,
            workspace.wire,
        )?
        .data
        .ok_or(Error::InvalidSource)?;
        self.structure_member_bytes(&data_member(&workspace, data)?)
            .ok_or(Error::InvalidSource)
    }

    /// Apply an exact-source-checked image-lifecycle patch.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when the patch was made from another
    /// package and typed source errors.
    pub fn apply_slide_image_lifecycle(&self, patch: &Patch) -> Result<Commit, Error> {
        let source = physical_source(self)?;
        if !patch.artifacts.authorizes_source(&source) {
            return Err(Error::PatchConflict);
        }
        let target = patch.artifacts.target();
        let package = if patch.is_noop() {
            self.clone()
        } else {
            self.structure_reopen(target.to_vec())
                .map_err(map_read_error)?
        };
        Ok(Commit {
            package,
            patch: patch.clone(),
            diagnostics: Diagnostics {
                changed: !patch.is_noop(),
                operations: 0,
            },
        })
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn vertical_flip_adds_a_half_turn() {
        assert_eq!(vertical_flip_angle(None), 180.0);
        assert_eq!(vertical_flip_angle(Some(180.0)), 0.0);
        assert_eq!(vertical_flip_angle(Some(270.0)), 90.0);
    }
}
