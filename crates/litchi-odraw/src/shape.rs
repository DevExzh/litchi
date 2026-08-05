//! Typed, bounded views over OfficeArt shape containers.

use bitflags::bitflags;

use crate::prop::{Anchor, Id, Props};
use crate::{Container, Error, Limit, Limits, Record, RecordKind, Result};

bitflags! {
    /// Flags encoded by an OfficeArt `Sp` atom (`[MS-ODRAW]` section 2.2.40).
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
    pub struct Flags: u32 {
        /// Shape is a group.
        const GROUP = 0x0001;
        /// Shape is a child of a group.
        const CHILD = 0x0002;
        /// Shape is the topmost group.
        const PATRIARCH = 0x0004;
        /// Shape has been deleted.
        const DELETED = 0x0008;
        /// Shape represents an OLE object.
        const OLE_SHAPE = 0x0010;
        /// Shape has a valid master.
        const HAVE_MASTER = 0x0020;
        /// Shape is flipped horizontally.
        const FLIP_H = 0x0040;
        /// Shape is flipped vertically.
        const FLIP_V = 0x0080;
        /// Shape is a connector.
        const CONNECTOR = 0x0100;
        /// Shape has an anchor.
        const HAVE_ANCHOR = 0x0200;
        /// Shape is a background shape.
        const BACKGROUND = 0x0400;
        /// Shape has an explicit primitive type.
        const HAVE_SPT = 0x0800;
    }
}

/// A format-neutral OfficeArt shape family.
///
/// Host applications project `ClientData` and `ClientTextbox` records into
/// their own richer shape enums instead of adding host-specific variants here.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Kind {
    /// Rectangle primitive.
    Rectangle,
    /// Ellipse primitive.
    Ellipse,
    /// Text box primitive.
    TextBox,
    /// Line primitive.
    Line,
    /// Connector primitive.
    Connector,
    /// Callout primitive.
    Callout,
    /// Freeform or polygon geometry.
    Polygon,
    /// Shape group.
    Group,
    /// Table group identified by OfficeArt table properties.
    Table,
    /// Picture-frame primitive.
    Picture,
    /// Other recognized OfficeArt primitive.
    AutoShape,
    /// Unknown or extension primitive.
    Unknown,
}

/// A lossless native `MSOSPT` value.
///
/// This newtype prevents shape kinds from being mixed with record, property,
/// or object identifiers while preserving extension values exactly.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(transparent)]
pub struct Native(u16);

impl Native {
    /// Non-primitive/freeform geometry.
    pub const FREEFORM: Self = Self(0);
    /// Rectangle primitive.
    pub const RECTANGLE: Self = Self(1);
    /// Rounded-rectangle primitive.
    pub const ROUND_RECTANGLE: Self = Self(2);
    /// Ellipse primitive.
    pub const ELLIPSE: Self = Self(3);
    /// Diamond primitive.
    pub const DIAMOND: Self = Self(4);
    /// Line primitive.
    pub const LINE: Self = Self(20);
    /// Picture-frame primitive.
    pub const PICTURE: Self = Self(75);
    /// Text-box primitive.
    pub const TEXT_BOX: Self = Self(202);

    /// Preserves a native value read from a producer or specification extension.
    pub const fn from_raw(raw: u16) -> Self {
        Self(raw)
    }

    /// Returns the exact native wire value.
    pub const fn raw(self) -> u16 {
        self.0
    }
}

/// The coordinate-space bounds carried by an OfficeArt `FSPGR` atom.
///
/// The bounds define the coordinate system in which child-shape anchors are
/// expressed. The original `Spgr` record remains available through
/// [`Shape::meta`] so future or producer-specific records stay lossless.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Bounds {
    /// Left boundary of the group coordinate system.
    pub left: i32,
    /// Top boundary of the group coordinate system.
    pub top: i32,
    /// Right boundary of the group coordinate system.
    pub right: i32,
    /// Bottom boundary of the group coordinate system.
    pub bottom: i32,
}

impl Bounds {
    /// Creates group coordinate-space bounds from their four wire values.
    #[inline]
    pub const fn new(left: i32, top: i32, right: i32, bottom: i32) -> Self {
        Self {
            left,
            top,
            right,
            bottom,
        }
    }

    /// Returns the checked horizontal extent.
    #[inline]
    pub const fn width(&self) -> Option<i32> {
        self.right.checked_sub(self.left)
    }

    /// Returns the checked vertical extent.
    #[inline]
    pub const fn height(&self) -> Option<i32> {
        self.bottom.checked_sub(self.top)
    }

    /// Decodes one exact `[MS-ODRAW]` `FSPGR` record without copying its
    /// payload. The four fixed-width coordinates are decoded by value, while
    /// the record and any neighboring unknown records remain borrowed by the
    /// containing shape.
    pub fn from_record(record: &Record<'_>) -> Result<Self> {
        validate_atom(record, RecordKind::Spgr, 1, Some(0), 16)?;
        Ok(Self::new(
            coordinate(record.data(), 0)?,
            coordinate(record.data(), 4)?,
            coordinate(record.data(), 8)?,
            coordinate(record.data(), 12)?,
        ))
    }
}

/// A parsed OfficeArt shape borrowing all variable-length data from its input.
///
/// The type intentionally does not implement `Clone`: callers move shape trees
/// or borrow them through [`Shape::children`] rather than accidentally copying
/// every node and property table.
#[derive(Debug)]
pub struct Shape<'data> {
    kind: Kind,
    id: u32,
    native_kind: Native,
    flags: Flags,
    props: Props<'data>,
    anchor: Option<Anchor>,
    group_bounds: Option<Bounds>,
    children: Vec<Shape<'data>>,
    container: Container<'data>,
    meta: Container<'data>,
    client_data: Option<Record<'data>>,
    textbox: Option<Record<'data>>,
    client_anchor: Option<Record<'data>>,
}

impl<'data> Shape<'data> {
    /// Returns the format-neutral shape family.
    pub const fn kind(&self) -> Kind {
        self.kind
    }

    /// Returns the OfficeArt shape identifier.
    pub const fn id(&self) -> u32 {
        self.id
    }

    /// Returns the native `MSOSPT` value from the shape atom.
    pub const fn native_kind(&self) -> Native {
        self.native_kind
    }

    /// Returns the typed flags from the shape atom.
    pub const fn flags(&self) -> Flags {
        self.flags
    }

    /// Returns the primary shape property table.
    pub const fn props(&self) -> &Props<'data> {
        &self.props
    }

    /// Returns the format-neutral child anchor when present.
    ///
    /// `ClientAnchor` payloads are defined by the host application, so DOC,
    /// PPT, and XLS callers decode those through [`Shape::client_anchor`].
    pub const fn anchor(&self) -> Option<&Anchor> {
        self.anchor.as_ref()
    }

    /// Returns the group coordinate system used by child-shape anchors.
    pub const fn group_bounds(&self) -> Option<&Bounds> {
        self.group_bounds.as_ref()
    }

    /// Borrows child shapes without copying the tree.
    pub fn children(&self) -> &[Shape<'data>] {
        &self.children
    }

    /// Returns the outer shape or shape-group container.
    pub const fn container(&self) -> &Container<'data> {
        &self.container
    }

    /// Returns the shape container that owns this shape's atom and properties.
    ///
    /// For a group this is its first `SpContainer`; for other shapes it is the
    /// same container returned by [`Shape::container`].
    pub const fn meta(&self) -> &Container<'data> {
        &self.meta
    }

    /// Borrows the host `ClientData` record without interpreting it.
    pub const fn client_data(&self) -> Option<&Record<'data>> {
        self.client_data.as_ref()
    }

    /// Borrows the host `ClientTextbox` record without interpreting it.
    pub const fn textbox(&self) -> Option<&Record<'data>> {
        self.textbox.as_ref()
    }

    /// Borrows the host `ClientAnchor` record without interpreting it.
    pub const fn client_anchor(&self) -> Option<&Record<'data>> {
        self.client_anchor.as_ref()
    }
}

/// Parses the user-visible shapes in one OfficeArt drawing.
pub fn parse(data: &[u8]) -> Result<Vec<Shape<'_>>> {
    parse_with(data, Limits::default())
}

/// Parses user-visible shapes within explicit depth and record ceilings.
pub fn parse_with(data: &[u8], limits: Limits) -> Result<Vec<Shape<'_>>> {
    const MAX_SAFE_DEPTH: u16 = 64;

    if limits.max_depth > MAX_SAFE_DEPTH {
        return Err(Error::InvalidLimit {
            limit: Limit::Depth,
            maximum: u32::from(MAX_SAFE_DEPTH),
        });
    }
    if data.is_empty() {
        return Ok(Vec::new());
    }

    let (record, consumed) = Record::parse(data, 0)?;
    if consumed != data.len() {
        return Err(Error::TrailingData { offset: consumed });
    }
    let root = Container::try_new(record)?;
    let mut budget = Budget::new(limits);
    budget.visit()?;
    budget.depth(0)?;

    match root.record().kind() {
        RecordKind::SpgrContainer => root_group(&root, 0, &mut budget),
        RecordKind::DgContainer => drawing(&root, 0, &mut budget),
        RecordKind::SpContainer => Ok(vec![build_shape(
            root,
            Vec::new(),
            Role::Standalone,
            &mut budget,
        )?]),
        _ => Err(Error::MalformedShape {
            reason: "root is not a drawing or shape container",
        }),
    }
}

impl<'data> TryFrom<Record<'data>> for Shape<'data> {
    type Error = Error;

    /// Builds one typed shape from an `SpContainer` or `SpgrContainer`.
    fn try_from(record: Record<'data>) -> Result<Self> {
        let kind = record.kind();
        let container = Container::try_new(record)?;
        let mut budget = Budget::new(Limits::default());
        budget.visit()?;
        match kind {
            RecordKind::SpContainer => {
                build_shape(container, Vec::new(), Role::Standalone, &mut budget)
            },
            RecordKind::SpgrContainer => build_group(container, 0, Role::Standalone, &mut budget),
            _ => Err(Error::MalformedShape {
                reason: "record is not a shape container",
            }),
        }
    }
}

#[derive(Debug)]
struct Budget {
    limits: Limits,
    records: u32,
}

impl Budget {
    const fn new(limits: Limits) -> Self {
        Self { limits, records: 0 }
    }

    fn visit(&mut self) -> Result<()> {
        self.records = self.records.checked_add(1).ok_or(Error::LimitExceeded {
            limit: Limit::Records,
            maximum: self.limits.max_records,
        })?;
        if self.records > self.limits.max_records {
            return Err(Error::LimitExceeded {
                limit: Limit::Records,
                maximum: self.limits.max_records,
            });
        }
        Ok(())
    }

    fn depth(&self, depth: u16) -> Result<()> {
        if depth > self.limits.max_depth {
            return Err(Error::LimitExceeded {
                limit: Limit::Depth,
                maximum: u32::from(self.limits.max_depth),
            });
        }
        Ok(())
    }
}

fn drawing<'data>(
    container: &Container<'data>,
    depth: u16,
    budget: &mut Budget,
) -> Result<Vec<Shape<'data>>> {
    validate_container_header(container, RecordKind::DgContainer)?;
    budget.depth(depth)?;

    let mut dg = false;
    let mut root_group_seen = false;
    let mut shapes = Vec::new();
    for child in container.children() {
        let child = child?;
        budget.visit()?;
        match child.kind() {
            RecordKind::Dg => {
                if dg {
                    return Err(Error::MalformedShape {
                        reason: "drawing contains more than one Dg atom",
                    });
                }
                validate_atom(&child, RecordKind::Dg, 0, None, 8)?;
                dg = true;
            },
            RecordKind::SpContainer => {
                let shape =
                    build_shape(Container::try_new(child)?, Vec::new(), Role::Root, budget)?;
                if !shape.flags().contains(Flags::BACKGROUND) {
                    shapes.push(shape);
                }
            },
            RecordKind::SpgrContainer => {
                if root_group_seen {
                    return Err(Error::MalformedShape {
                        reason: "drawing contains more than one root shape group",
                    });
                }
                root_group_seen = true;
                shapes.extend(root_group(
                    &Container::try_new(child)?,
                    next_depth(depth)?,
                    budget,
                )?);
            },
            RecordKind::SolverContainer => {
                validate_container_record(&child, RecordKind::SolverContainer)?;
            },
            RecordKind::Unknown(_) => {},
            _ => {
                return Err(Error::MalformedShape {
                    reason: "drawing contains an invalid direct child record",
                });
            },
        }
    }
    if !dg {
        return Err(Error::MalformedShape {
            reason: "drawing has no Dg atom",
        });
    }
    Ok(shapes)
}

fn root_group<'data>(
    container: &Container<'data>,
    depth: u16,
    budget: &mut Budget,
) -> Result<Vec<Shape<'data>>> {
    validate_container_header(container, RecordKind::SpgrContainer)?;
    budget.depth(depth)?;
    let mut children = container.children();
    let header = children.next().ok_or(Error::MalformedShape {
        reason: "shape-group container has no group header",
    })??;
    budget.visit()?;
    if header.kind() != RecordKind::SpContainer {
        return Err(Error::MalformedShape {
            reason: "shape-group container does not start with SpContainer",
        });
    }
    let header = Container::try_new(header)?;
    let meta = scan_meta(&header, budget)?;
    validate_meta(&meta, true, Role::Patriarch)?;

    let mut shapes = Vec::new();
    for child in children {
        let child = child?;
        budget.visit()?;
        match child.kind() {
            RecordKind::SpContainer => {
                let shape =
                    build_shape(Container::try_new(child)?, Vec::new(), Role::Root, budget)?;
                if !shape.flags().contains(Flags::BACKGROUND) {
                    shapes.push(shape);
                }
            },
            RecordKind::SpgrContainer => shapes.push(build_group(
                Container::try_new(child)?,
                next_depth(depth)?,
                Role::Root,
                budget,
            )?),
            _ => {
                return Err(Error::MalformedShape {
                    reason: "shape-group container has a non-shape child",
                });
            },
        }
    }
    Ok(shapes)
}

fn build_group<'data>(
    container: Container<'data>,
    depth: u16,
    role: Role,
    budget: &mut Budget,
) -> Result<Shape<'data>> {
    validate_container_header(&container, RecordKind::SpgrContainer)?;
    budget.depth(depth)?;
    let mut records = container.children();
    let header = records.next().ok_or(Error::MalformedShape {
        reason: "shape-group container has no group header",
    })??;
    budget.visit()?;
    if header.kind() != RecordKind::SpContainer {
        return Err(Error::MalformedShape {
            reason: "shape-group container does not start with SpContainer",
        });
    }
    let meta = Container::try_new(header)?;

    let mut children = Vec::new();
    for child in records {
        let child = child?;
        budget.visit()?;
        match child.kind() {
            RecordKind::SpContainer => {
                children.push(build_shape(
                    Container::try_new(child)?,
                    Vec::new(),
                    Role::Member,
                    budget,
                )?);
            },
            RecordKind::SpgrContainer => children.push(build_group(
                Container::try_new(child)?,
                next_depth(depth)?,
                Role::Member,
                budget,
            )?),
            _ => {
                return Err(Error::MalformedShape {
                    reason: "shape-group container has a non-shape child",
                });
            },
        }
    }

    build(container, meta, children, true, role, budget)
}

fn build_shape<'data>(
    container: Container<'data>,
    children: Vec<Shape<'data>>,
    role: Role,
    budget: &mut Budget,
) -> Result<Shape<'data>> {
    build(container.clone(), container, children, false, role, budget)
}

fn build<'data>(
    container: Container<'data>,
    meta: Container<'data>,
    children: Vec<Shape<'data>>,
    group: bool,
    role: Role,
    budget: &mut Budget,
) -> Result<Shape<'data>> {
    let mut records = scan_meta(&meta, budget)?;
    let (id, native_kind, flags, anchor) = validate_meta(&records, group, role)?;
    let group_bounds = records.spgr.as_ref().map(Bounds::from_record).transpose()?;
    let props = records.primary.take().unwrap_or_else(Props::new);
    let kind = if group {
        if records
            .tertiary
            .as_ref()
            .and_then(|props| props.get_int(Id::GroupTableProperties))
            .is_some_and(|value| value & 1 != 0)
        {
            Kind::Table
        } else {
            Kind::Group
        }
    } else {
        detect(native_kind, &props)
    };

    Ok(Shape {
        kind,
        id,
        native_kind,
        flags,
        props,
        anchor,
        group_bounds,
        children,
        container,
        meta,
        client_data: records.client_data,
        textbox: records.textbox,
        client_anchor: records.client_anchor,
    })
}

#[derive(Debug, Clone, Copy)]
enum Role {
    Patriarch,
    Root,
    Member,
    Standalone,
}

#[derive(Debug)]
struct Meta<'data> {
    sp: Option<Record<'data>>,
    spgr: Option<Record<'data>>,
    primary: Option<Props<'data>>,
    secondary: Option<Props<'data>>,
    tertiary: Option<Props<'data>>,
    child_anchor: Option<Record<'data>>,
    client_anchor: Option<Record<'data>>,
    client_data: Option<Record<'data>>,
    textbox: Option<Record<'data>>,
}

impl Meta<'_> {
    const fn new() -> Self {
        Self {
            sp: None,
            spgr: None,
            primary: None,
            secondary: None,
            tertiary: None,
            child_anchor: None,
            client_anchor: None,
            client_data: None,
            textbox: None,
        }
    }
}

fn scan_meta<'data>(container: &Container<'data>, budget: &mut Budget) -> Result<Meta<'data>> {
    validate_container_header(container, RecordKind::SpContainer)?;
    let mut meta = Meta::new();
    for child in container.children() {
        let child = child?;
        budget.visit()?;
        match child.kind() {
            RecordKind::Sp => {
                validate_atom(&child, RecordKind::Sp, 2, None, 8)?;
                insert(&mut meta.sp, child)?;
            },
            RecordKind::Spgr => {
                validate_atom(&child, RecordKind::Spgr, 1, Some(0), 16)?;
                insert(&mut meta.spgr, child)?;
            },
            RecordKind::Opt => {
                let props = Props::parse(&child)?;
                insert(&mut meta.primary, props)?;
            },
            RecordKind::SecondaryOpt => {
                let props = Props::parse(&child)?;
                insert(&mut meta.secondary, props)?;
            },
            RecordKind::TertiaryOpt => {
                let props = Props::parse(&child)?;
                insert(&mut meta.tertiary, props)?;
            },
            RecordKind::ChildAnchor => {
                validate_atom(&child, RecordKind::ChildAnchor, 0, Some(0), 16)?;
                insert(&mut meta.child_anchor, child)?;
            },
            RecordKind::ClientAnchor => {
                validate_atom_kind(&child, RecordKind::ClientAnchor, 0, Some(0))?;
                insert(&mut meta.client_anchor, child)?;
            },
            RecordKind::ClientData => insert(&mut meta.client_data, child)?,
            RecordKind::ClientTextbox => insert(&mut meta.textbox, child)?,
            RecordKind::Unknown(_) => {},
            _ => {
                return Err(Error::MalformedShape {
                    reason: "SpContainer contains an invalid child record",
                });
            },
        }
    }
    Ok(meta)
}

fn validate_meta(
    meta: &Meta<'_>,
    group: bool,
    role: Role,
) -> Result<(u32, Native, Flags, Option<Anchor>)> {
    let sp = meta.sp.as_ref().ok_or(Error::MalformedShape {
        reason: "SpContainer has no shape atom",
    })?;
    let data: &[u8; 8] = sp.data().try_into().map_err(|_| Error::MalformedShape {
        reason: "shape atom payload is not eight bytes",
    })?;
    let id = u32::from_le_bytes(data[..4].try_into().map_err(|_| Error::MalformedShape {
        reason: "shape identifier is truncated",
    })?);
    let raw_flags =
        u32::from_le_bytes(data[4..].try_into().map_err(|_| Error::MalformedShape {
            reason: "shape flags are truncated",
        })?);
    let native = Native::from_raw(sp.instance());
    let flags = Flags::from_bits_retain(raw_flags);

    if flags.contains(Flags::GROUP) != group {
        return Err(Error::MalformedShape {
            reason: "shape GROUP flag disagrees with its container topology",
        });
    }
    if group != meta.spgr.is_some() {
        return Err(Error::MalformedShape {
            reason: "group shape must contain exactly one Spgr atom",
        });
    }
    if group && native != Native::FREEFORM {
        return Err(Error::MalformedShape {
            reason: "group shape must use the non-primitive native kind",
        });
    }

    let patriarch = matches!(role, Role::Patriarch);
    if flags.contains(Flags::PATRIARCH) != patriarch && !matches!(role, Role::Standalone) {
        return Err(Error::MalformedShape {
            reason: "shape PATRIARCH flag disagrees with its container topology",
        });
    }
    let expected_child = match role {
        Role::Patriarch | Role::Root => Some(false),
        Role::Member => Some(true),
        Role::Standalone => None,
    };
    if expected_child.is_some_and(|expected| flags.contains(Flags::CHILD) != expected) {
        return Err(Error::MalformedShape {
            reason: "shape CHILD flag disagrees with its group membership",
        });
    }

    if meta.child_anchor.is_some() && meta.client_anchor.is_some() {
        return Err(Error::MalformedShape {
            reason: "shape contains both child and host anchors",
        });
    }
    let has_anchor = meta.child_anchor.is_some() || meta.client_anchor.is_some();
    // Word emits a direct background shape whose fHaveAnchor bit is set even
    // though its host-owned anchor is omitted. Keep the structural checks for
    // every user shape, but accept this producer-specific, non-visible sentinel.
    let word_background_sentinel =
        flags.contains(Flags::BACKGROUND | Flags::HAVE_ANCHOR) && !has_anchor;
    if flags.contains(Flags::HAVE_ANCHOR) != has_anchor && !word_background_sentinel {
        return Err(Error::MalformedShape {
            reason: "shape HAVE_ANCHOR flag disagrees with its anchor records",
        });
    }
    if flags.contains(Flags::CHILD) {
        if meta.client_anchor.is_some() {
            return Err(Error::MalformedShape {
                reason: "group child uses a host ClientAnchor",
            });
        }
    } else if meta.child_anchor.is_some() {
        return Err(Error::MalformedShape {
            reason: "non-child shape uses a ChildAnchor",
        });
    }

    let anchor = meta
        .child_anchor
        .as_ref()
        .map(|record| {
            Anchor::from_child_anchor(record).ok_or(Error::MalformedShape {
                reason: "child anchor payload is not sixteen bytes",
            })
        })
        .transpose()?;

    Ok((id, native, flags, anchor))
}

fn detect(native: Native, props: &Props<'_>) -> Kind {
    match native.raw() {
        0 if props.has(Id::Vertices) => Kind::Polygon,
        0 => Kind::AutoShape,
        1 => Kind::Rectangle,
        3 => Kind::Ellipse,
        20 => Kind::Line,
        32..=40 => Kind::Connector,
        41..=52 | 61..=63 | 106 | 178..=189 => Kind::Callout,
        75 => Kind::Picture,
        202 => Kind::TextBox,
        2..=201 => Kind::AutoShape,
        _ => Kind::Unknown,
    }
}

fn insert<T>(slot: &mut Option<T>, value: T) -> Result<()> {
    if slot.replace(value).is_some() {
        return Err(Error::MalformedShape {
            reason: "shape container contains a duplicate singleton record",
        });
    }
    Ok(())
}

fn validate_container_header(container: &Container<'_>, kind: RecordKind) -> Result<()> {
    validate_container_record(container.record(), kind)
}

fn validate_container_record(record: &Record<'_>, kind: RecordKind) -> Result<()> {
    if record.kind() != kind
        || record.raw_kind() != kind.raw()
        || record.version() != 0x0F
        || record.instance() != 0
    {
        return Err(Error::MalformedShape {
            reason: "OfficeArt container header is invalid",
        });
    }
    Ok(())
}

fn validate_atom(
    record: &Record<'_>,
    kind: RecordKind,
    version: u8,
    instance: Option<u16>,
    len: u32,
) -> Result<()> {
    validate_atom_kind(record, kind, version, instance)?;
    if record.len() != len || usize::try_from(len).ok() != Some(record.data().len()) {
        return Err(Error::MalformedShape {
            reason: "OfficeArt atom payload length is invalid",
        });
    }
    Ok(())
}

fn validate_atom_kind(
    record: &Record<'_>,
    kind: RecordKind,
    version: u8,
    instance: Option<u16>,
) -> Result<()> {
    if record.kind() != kind
        || record.raw_kind() != kind.raw()
        || record.version() != version
        || instance.is_some_and(|expected| record.instance() != expected)
    {
        return Err(Error::MalformedShape {
            reason: "OfficeArt atom header is invalid",
        });
    }
    Ok(())
}

fn next_depth(depth: u16) -> Result<u16> {
    depth.checked_add(1).ok_or(Error::LimitExceeded {
        limit: Limit::Depth,
        maximum: u32::from(u16::MAX),
    })
}

fn coordinate(data: &[u8], offset: usize) -> Result<i32> {
    let end = offset.checked_add(4).ok_or(Error::ArithmeticOverflow {
        context: "group coordinate extent",
    })?;
    let bytes = data.get(offset..end).ok_or(Error::MalformedShape {
        reason: "group coordinate atom payload is truncated",
    })?;
    let bytes: [u8; 4] = bytes.try_into().map_err(|_| Error::MalformedShape {
        reason: "group coordinate atom payload is truncated",
    })?;
    Ok(i32::from_le_bytes(bytes))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::write::{self, Atom, Container as OutContainer, ShapeBuilder};

    fn shape(kind: Native, id: u32, child: bool) -> Vec<u8> {
        let mut body = Vec::new();
        let mut flags = Flags::HAVE_ANCHOR | Flags::HAVE_SPT;
        if child {
            flags |= Flags::CHILD;
        }
        ShapeBuilder::new(kind, id)
            .with_flags(flags)
            .write(&mut body)
            .expect("write shape atom");
        if child {
            write::child_anchor(&mut body, 10, 20, 110, 70).expect("write child anchor");
        } else {
            let mut payload = Vec::new();
            for coordinate in [10_i32, 20, 110, 70] {
                payload.extend_from_slice(&coordinate.to_le_bytes());
            }
            write::atom(&mut body, 0, Atom::ClientAnchor, &payload)
                .expect("write opaque host anchor");
        }
        let mut record = Vec::new();
        write::container(&mut record, 0, OutContainer::Sp, &body).expect("write shape container");
        record
    }

    fn patriarch(id: u32) -> Vec<u8> {
        let mut body = Vec::new();
        write::spgr(&mut body, 0, 0, 0, 0).expect("write patriarch bounds");
        ShapeBuilder::new(Native::FREEFORM, id)
            .with_flags(Flags::GROUP | Flags::PATRIARCH)
            .write(&mut body)
            .expect("write patriarch shape atom");
        let mut record = Vec::new();
        write::container(&mut record, 0, OutContainer::Sp, &body)
            .expect("write patriarch container");
        record
    }

    fn background(id: u32) -> Vec<u8> {
        let mut body = Vec::new();
        ShapeBuilder::new(Native::RECTANGLE, id)
            .with_flags(Flags::BACKGROUND | Flags::HAVE_ANCHOR | Flags::HAVE_SPT)
            .write(&mut body)
            .expect("write background shape atom");
        let mut record = Vec::new();
        write::container(&mut record, 0, OutContainer::Sp, &body)
            .expect("write background container");
        record
    }

    fn drawing() -> Vec<u8> {
        let patriarch = patriarch(1);
        let rectangle = shape(Native::RECTANGLE, 2, false);

        let mut group_header_body = Vec::new();
        write::spgr(&mut group_header_body, 0, 0, 1000, 500).expect("write group bounds");
        let future = Atom::unknown(0xF123, 0).expect("future atom kind");
        write::atom(&mut group_header_body, 0, future, &[0xAA, 0xBB]).expect("write future atom");
        ShapeBuilder::new(Native::FREEFORM, 3)
            .with_flags(Flags::GROUP | Flags::HAVE_ANCHOR)
            .write(&mut group_header_body)
            .expect("write group shape");
        let mut group_anchor = Vec::new();
        for coordinate in [100_i32, 200, 500, 400] {
            group_anchor.extend_from_slice(&coordinate.to_le_bytes());
        }
        write::atom(&mut group_header_body, 0, Atom::ClientAnchor, &group_anchor)
            .expect("write group anchor");
        let mut group_header = Vec::new();
        write::container(&mut group_header, 0, OutContainer::Sp, &group_header_body)
            .expect("write group header");

        let mut nested_body = group_header;
        nested_body.extend_from_slice(&shape(Native::ELLIPSE, 4, true));
        let mut nested = Vec::new();
        write::container(&mut nested, 0, OutContainer::Spgr, &nested_body)
            .expect("write nested group");

        let mut root_body = patriarch;
        root_body.extend_from_slice(&rectangle);
        root_body.extend_from_slice(&nested);
        let mut root = Vec::new();
        write::container(&mut root, 0, OutContainer::Spgr, &root_body).expect("write root group");

        let mut drawing_body = Vec::new();
        write::dg(&mut drawing_body, 5, 5).expect("write drawing atom");
        drawing_body.extend_from_slice(&root);
        drawing_body.extend_from_slice(&background(5));
        let mut bytes = Vec::new();
        write::container(&mut bytes, 0, OutContainer::Dg, &drawing_body).expect("write drawing");
        bytes
    }

    #[test]
    fn hides_root_patriarch_and_preserves_nested_group() {
        let bytes = drawing();
        let shapes = parse(&bytes).expect("parse drawing");

        assert_eq!(shapes.len(), 2);
        assert_eq!(shapes[0].kind(), Kind::Rectangle);
        assert_eq!(shapes[0].id(), 2);
        assert!(shapes[0].anchor().is_none());
        assert!(shapes[0].client_anchor().is_some());
        assert_eq!(shapes[1].kind(), Kind::Group);
        assert_eq!(shapes[1].id(), 3);
        assert_eq!(
            shapes[1].group_bounds().copied(),
            Some(Bounds::new(0, 0, 1000, 500))
        );
        assert_eq!(shapes[1].children()[0].kind(), Kind::Ellipse);
        assert_eq!(shapes[1].children()[0].id(), 4);
        assert!(
            shapes
                .iter()
                .all(|shape| !shape.flags().contains(Flags::BACKGROUND))
        );
    }

    #[test]
    fn rejects_a_missing_anchor_for_a_user_shape() {
        let mut body = Vec::new();
        ShapeBuilder::new(Native::RECTANGLE, 1)
            .with_flags(Flags::HAVE_ANCHOR | Flags::HAVE_SPT)
            .write(&mut body)
            .expect("write shape atom");
        let mut bytes = Vec::new();
        write::container(&mut bytes, 0, OutContainer::Sp, &body).expect("write shape container");

        assert!(matches!(
            parse(&bytes),
            Err(Error::MalformedShape {
                reason: "shape HAVE_ANCHOR flag disagrees with its anchor records",
            })
        ));
    }

    #[test]
    fn traversal_limits_are_enforced() {
        let error = parse_with(
            &drawing(),
            Limits {
                max_depth: 0,
                max_records: 1_000,
            },
        )
        .expect_err("nested group exceeds depth zero");

        assert!(matches!(
            error,
            Error::LimitExceeded {
                limit: Limit::Depth,
                ..
            }
        ));
    }

    #[test]
    fn rejects_trailing_root_bytes() {
        let mut bytes = drawing();
        bytes.push(0);

        assert!(matches!(parse(&bytes), Err(Error::TrailingData { .. })));
    }

    #[test]
    fn rejects_an_unsafe_recursive_depth_limit() {
        assert!(matches!(
            parse_with(
                &drawing(),
                Limits {
                    max_depth: 65,
                    max_records: 1_000,
                },
            ),
            Err(Error::InvalidLimit {
                limit: Limit::Depth,
                maximum: 64,
            })
        ));
    }

    #[test]
    fn metadata_records_consume_the_record_budget() {
        assert!(matches!(
            parse_with(
                &drawing(),
                Limits {
                    max_depth: 64,
                    max_records: 5,
                },
            ),
            Err(Error::LimitExceeded {
                limit: Limit::Records,
                maximum: 5,
            })
        ));
    }

    #[test]
    fn rejects_a_group_coordinate_atom_with_the_wrong_length() {
        let data = [0_u8; 12];
        let record = Record::from_parts(RecordKind::Spgr, 1, 0, &data).expect("test record");

        assert!(matches!(
            Bounds::from_record(&record),
            Err(Error::MalformedShape {
                reason: "OfficeArt atom payload length is invalid",
            })
        ));
    }

    #[test]
    fn typed_group_bounds_do_not_drop_future_records() {
        let bytes = drawing();
        let shapes = parse(&bytes).expect("parse drawing");
        let group = &shapes[1];
        let extension = group
            .meta()
            .find(RecordKind::Unknown(0xF123))
            .expect("scan unknown record")
            .expect("future record");

        assert_eq!(extension.data(), &[0xAA, 0xBB]);
        assert!(extension.data_offset(&bytes).is_some());
    }
}
