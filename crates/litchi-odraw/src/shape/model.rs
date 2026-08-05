//! Borrowed, format-neutral OfficeArt shape objects.

use bitflags::bitflags;

use crate::prop::{Anchor, Props};
use crate::{Container, Record};

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

impl From<u32> for Flags {
    fn from(bits: u32) -> Self {
        Self::from_bits_retain(bits)
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

impl From<u16> for Native {
    fn from(raw: u16) -> Self {
        Self::from_raw(raw)
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

    pub(crate) fn from_parts(
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
    ) -> Self {
        Self {
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
            client_data,
            textbox,
            client_anchor,
        }
    }
}
