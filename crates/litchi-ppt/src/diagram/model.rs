//! Contextual diagram inventory values.

use crate::animation::diagram_build;
use crate::odraw::{Drawing, Shape};
use crate::package::Result;
use litchi_odraw::{Record, RecordKind};

/// The stable identity of one diagram build on a slide.
///
/// [MS-PPT] requires the pair of `buildId` and `shapeIdRef` to be unique for
/// all builds on a slide.  Keeping the pair together prevents callers from
/// accidentally treating either native integer as a global identifier.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Id {
    build_id: u32,
    shape_id: u32,
}

impl Id {
    /// Creates a checked semantic identity from the two MS-PPT fields.
    pub const fn new(build_id: u32, shape_id: u32) -> Self {
        Self { build_id, shape_id }
    }

    /// Returns the PowerPoint build identifier.
    pub const fn build_id(self) -> u32 {
        self.build_id
    }

    /// Returns the referenced OfficeArt shape identifier.
    pub const fn shape_id(self) -> u32 {
        self.shape_id
    }
}

/// Diagram-specific view of the shared PowerPoint build atom.
///
/// The underlying fixed-width record is owned by
/// [`crate::animation::diagram_build`].  This wrapper only supplies concise
/// diagram-context accessors and therefore does not introduce a second wire
/// parser or a second copy of the record fields.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Build {
    record: diagram_build::Container,
}

impl Build {
    pub(super) const fn new(record: diagram_build::Container) -> Self {
        Self { record }
    }

    /// Returns the diagram identity represented by this build.
    pub const fn id(self) -> Id {
        Id::new(
            self.record.build().build_id,
            self.record.build().shape_id_ref,
        )
    }

    /// Returns the PowerPoint build identifier.
    pub const fn build_id(self) -> u32 {
        self.record.build().build_id
    }

    /// Returns the target OfficeArt shape identifier.
    pub const fn shape_id(self) -> u32 {
        self.record.build().shape_id_ref
    }

    /// Whether the build is expanded into timing nodes.
    pub const fn expanded(self) -> bool {
        self.record.build().expanded
    }

    /// Whether the build is expanded in the authoring UI.
    pub const fn ui_expanded(self) -> bool {
        self.record.build().ui_expanded
    }

    /// Returns the two undefined bytes retained by the source BuildAtom.
    pub const fn reserved(self) -> [u8; 2] {
        self.record.build().reserved()
    }

    /// Returns the MS-PPT diagram build mode, retaining unknown values.
    pub const fn mode(self) -> diagram_build::BuildType {
        self.record.atom().build_type
    }

    /// Returns the lossless fixed-width owner for advanced callers.
    pub const fn record(self) -> diagram_build::Container {
        self.record
    }
}

/// A lightweight reference to an OfficeArt shape associated with a diagram.
///
/// The reference contains only the native shape identifier.  Resolution is
/// performed against the inventory's borrowed ODraw tree, so a diagram entry
/// does not clone or flatten shape objects.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ShapeRef {
    id: u32,
}

impl ShapeRef {
    pub(super) const fn new(id: u32) -> Self {
        Self { id }
    }

    /// Returns the native OfficeArt shape identifier.
    pub const fn id(self) -> u32 {
        self.id
    }

    /// Resolves this reference in a parsed drawing without allocating.
    pub fn resolve<'drawing, 'data>(
        self,
        drawing: &'drawing Drawing<'data>,
    ) -> Option<&'drawing Shape<'data>> {
        find_shape(drawing.shapes(), self.id)
    }
}

/// The role of one borrowed, uninterpreted OfficeArt record.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PayloadKind {
    /// The shape's own `SpContainer` record.
    Shape,
    /// The enclosing `SpgrContainer` record for a group shape.
    Group,
    /// Host-specific `ClientData` retained without interpretation.
    ClientData,
    /// Host-specific `ClientTextbox` retained without interpretation.
    Textbox,
    /// Host-specific `ClientAnchor` retained without interpretation.
    Anchor,
}

/// An inert OfficeArt payload reference.
///
/// The record remains borrowed from the caller's drawing bytes and can be
/// inspected or handed to a lower-level owner.  This facade never interprets
/// diagram layout, relation tables, or vendor-specific payload semantics.
#[derive(Debug, Clone)]
pub struct Payload<'data> {
    shape_id: u32,
    kind: PayloadKind,
    record: Record<'data>,
}

impl<'data> Payload<'data> {
    pub(super) const fn new(shape_id: u32, kind: PayloadKind, record: Record<'data>) -> Self {
        Self {
            shape_id,
            kind,
            record,
        }
    }

    /// Returns the shape to which this payload belongs.
    pub const fn shape_id(&self) -> u32 {
        self.shape_id
    }

    /// Returns the contextual payload role.
    pub const fn kind(&self) -> PayloadKind {
        self.kind
    }

    /// Returns the lossless OfficeArt record handle.
    pub const fn record(&self) -> &Record<'data> {
        &self.record
    }

    /// Returns the record's raw OfficeArt kind.
    pub const fn raw_kind(&self) -> u16 {
        self.record.raw_kind()
    }

    /// Returns the typed OfficeArt kind, including an unknown fallback.
    pub const fn record_kind(&self) -> RecordKind {
        self.record.kind()
    }

    /// Returns the record body without interpreting it.
    pub const fn bytes(&self) -> &'data [u8] {
        self.record.data()
    }
}

/// One native diagram inventory entry.
#[derive(Debug)]
pub struct Diagram<'data> {
    id: Id,
    build: Build,
    root: ShapeRef,
    shapes: Vec<ShapeRef>,
    payloads: Vec<Payload<'data>>,
}

impl<'data> Diagram<'data> {
    pub(super) fn new(build: Build, shapes: Vec<ShapeRef>, payloads: Vec<Payload<'data>>) -> Self {
        let root = ShapeRef::new(build.shape_id());
        Self {
            id: build.id(),
            build,
            root,
            shapes,
            payloads,
        }
    }

    /// Returns the stable `(build_id, shape_id)` identity.
    pub const fn id(&self) -> Id {
        self.id
    }

    /// Returns the typed build metadata.
    pub const fn build(&self) -> Build {
        self.build
    }

    /// Returns the shape at which the diagram build is rooted.
    pub const fn root(&self) -> ShapeRef {
        self.root
    }

    /// Returns the root and all nested OfficeArt shapes in source order.
    pub fn shapes(&self) -> &[ShapeRef] {
        &self.shapes
    }

    /// Returns inert payload references in associated-shape/source order.
    pub fn payloads(&self) -> &[Payload<'data>] {
        &self.payloads
    }
}

/// Resource ceilings for one native diagram inventory.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// Maximum diagram builds retained from one build list.
    pub max_diagrams: usize,
    /// Maximum associated shapes retained for one diagram.
    pub max_shapes_per_diagram: usize,
    /// Maximum inert payload references retained for one diagram.
    pub max_payloads_per_diagram: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_diagrams: 1024,
            max_shapes_per_diagram: 1_000_000,
            max_payloads_per_diagram: 4_000_000,
        }
    }
}

/// Read-only native diagram inventory for one PPT drawing/build-list pair.
#[derive(Debug)]
pub struct Inventory<'data> {
    drawing: Drawing<'data>,
    diagrams: Vec<Diagram<'data>>,
}

impl<'data> Inventory<'data> {
    pub(super) fn new(drawing: Drawing<'data>, diagrams: Vec<Diagram<'data>>) -> Self {
        Self { drawing, diagrams }
    }

    /// Parses a build list together with its associated PPDrawing payload.
    pub fn parse(build_list: &crate::records::Record, drawing: &'data [u8]) -> Result<Self> {
        super::codec::parse(build_list, drawing)
    }

    /// Parses with explicit resource ceilings.
    pub fn parse_with_limits(
        build_list: &crate::records::Record,
        drawing: &'data [u8],
        limits: Limits,
    ) -> Result<Self> {
        super::codec::parse_with_limits(build_list, drawing, limits)
    }

    /// Returns the full borrowed ODraw projection used by this inventory.
    pub const fn drawing(&self) -> &Drawing<'data> {
        &self.drawing
    }

    /// Returns diagrams in BuildList source order.
    pub fn diagrams(&self) -> &[Diagram<'data>] {
        &self.diagrams
    }

    /// Returns the number of native diagram builds.
    pub fn len(&self) -> usize {
        self.diagrams.len()
    }

    /// Whether the build list contains no diagram builds.
    pub fn is_empty(&self) -> bool {
        self.diagrams.is_empty()
    }

    /// Finds a diagram by its checked build identity.
    pub fn get(&self, id: Id) -> Option<&Diagram<'data>> {
        self.diagrams.iter().find(|diagram| diagram.id() == id)
    }

    /// Resolves an associated shape reference against the retained ODraw tree.
    pub fn shape<'inventory>(
        &'inventory self,
        reference: ShapeRef,
    ) -> Option<&'inventory Shape<'data>> {
        reference.resolve(&self.drawing)
    }
}

fn find_shape<'drawing, 'data>(
    shapes: &'drawing [Shape<'data>],
    id: u32,
) -> Option<&'drawing Shape<'data>> {
    for shape in shapes {
        if shape.id() == id {
            return Some(shape);
        }
        if let Some(found) = find_shape(shape.children(), id) {
            return Some(found);
        }
    }
    None
}
