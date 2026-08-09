//! Typed Word shape models projected from `OfficeArt`.

use litchi_odraw::{Container, Record, RecordKind};

pub use litchi_odraw::shape::{Bounds, Flags, Kind, Native};
use litchi_odraw::{
    prop::{Id as OfficeArtPropertyId, Prop, Props, Value},
    shape::Shape as OfficeArtShape,
};

use std::io;

/// The stable identity carried by an `OfficeArt` `FSP` shape atom.
///
/// The value is deliberately kept separate from property and record
/// identifiers. It remains lossless for producer-defined identifiers while
/// giving callers a contextual selector for a projected shape tree.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
#[repr(transparent)]
pub struct ShapeId(u32);

impl ShapeId {
    /// Wraps the exact `spid` value from an `OfficeArt` shape atom.
    #[inline]
    #[must_use]
    pub const fn from_raw(raw: u32) -> Self {
        Self(raw)
    }

    /// Returns the exact `spid` value.
    #[inline]
    #[must_use]
    pub const fn raw(self) -> u32 {
        self.0
    }
}

impl From<u32> for ShapeId {
    fn from(value: u32) -> Self {
        Self::from_raw(value)
    }
}

impl From<ShapeId> for u32 {
    fn from(value: ShapeId) -> Self {
        value.raw()
    }
}

/// The four signed coordinates of an `OfficeArt` `ChildAnchor` record.
///
/// A child anchor is expressed in the coordinate space supplied by the
/// containing group's [`Bounds`]. Host-owned `ClientAnchor` bytes are kept in
/// [`ClientAnchor`] instead of being guessed into this coordinate system.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Anchor {
    /// Left coordinate in the containing group coordinate space.
    pub left: i32,
    /// Top coordinate in the containing group coordinate space.
    pub top: i32,
    /// Right coordinate in the containing group coordinate space.
    pub right: i32,
    /// Bottom coordinate in the containing group coordinate space.
    pub bottom: i32,
}

impl Anchor {
    /// Creates an anchor from its four signed `OfficeArt` coordinates.
    #[inline]
    #[must_use]
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
    #[must_use]
    pub const fn width(&self) -> Option<i32> {
        self.right.checked_sub(self.left)
    }

    /// Returns the checked vertical extent.
    #[inline]
    #[must_use]
    pub const fn height(&self) -> Option<i32> {
        self.bottom.checked_sub(self.top)
    }
}

impl From<litchi_odraw::prop::Anchor> for Anchor {
    fn from(anchor: litchi_odraw::prop::Anchor) -> Self {
        Self::new(anchor.left, anchor.top, anchor.right, anchor.bottom)
    }
}

/// An owned Word host-anchor record.
///
/// `[MS-ODRAW]` leaves `ClientAnchor` interpretation to the host format. DOC
/// commonly stores a four-byte index into `PlcfSpa`, while inline drawings may
/// use an empty payload. The complete record is retained so producer-specific
/// payloads remain lossless without pretending to be universal coordinates.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ClientAnchor {
    bytes: Box<[u8]>,
}

impl ClientAnchor {
    fn from_record(record: &Record<'_>) -> Self {
        Self {
            bytes: record_bytes(record),
        }
    }

    /// Returns the exact `OfficeArt` record bytes, including its header.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Returns the host-owned payload without the `OfficeArt` record header.
    #[must_use]
    pub fn payload(&self) -> &[u8] {
        &self.bytes[8..]
    }

    /// Returns DOC's common four-byte `PlcfSpa` index representation.
    ///
    /// Other host payload lengths remain available through [`Self::payload`].
    pub fn index(&self) -> Option<u32> {
        self.payload().try_into().ok().map(u32::from_le_bytes)
    }

    /// Returns whether the host record has no payload bytes.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.payload().is_empty()
    }
}

/// The property-table owner of an unknown `OfficeArt` property.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PropertyTable {
    /// The primary `OfficeArtFOPT` table.
    Primary,
    /// The secondary `OfficeArtSecondaryFOPT` table.
    Secondary,
    /// The tertiary `OfficeArtTertiaryFOPT` table.
    Tertiary,
}

/// One unknown `FOPTE` entry retained without a lossy property projection.
///
/// The six-byte descriptor and, for complex properties, its exact complex
/// payload are owned together. This permits snapshot use after the source OLE
/// stream is released and keeps future `[MS-ODRAW]` extensions inspectable.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UnknownProperty {
    table: PropertyTable,
    bytes: Box<[u8]>,
}

impl UnknownProperty {
    fn from_prop(table: PropertyTable, prop: &Prop<'_>) -> Option<Self> {
        if !matches!(prop.id(), OfficeArtPropertyId::Unknown(_)) {
            return None;
        }

        let complex = match prop.value() {
            Value::Simple(_) => None,
            Value::Complex(data) => Some(*data),
            Value::Array(array) => Some(array.raw_data()),
        };
        let mut bytes = Vec::with_capacity(6 + complex.map_or(0, <[u8]>::len));
        bytes.extend_from_slice(&prop.raw_opid().to_le_bytes());
        bytes.extend_from_slice(&prop.raw_value().to_le_bytes());
        if let Some(data) = complex {
            bytes.extend_from_slice(data);
        }

        Some(Self {
            table,
            bytes: bytes.into_boxed_slice(),
        })
    }

    /// Returns the property-table owner.
    #[must_use]
    pub const fn table(&self) -> PropertyTable {
        self.table
    }

    /// Returns the unflagged property identifier.
    #[must_use]
    pub fn raw_id(&self) -> u16 {
        self.raw_opid() & 0x3FFF
    }

    /// Returns the exact `opid`, including `fBid` and `fComplex` flags.
    #[must_use]
    pub fn raw_opid(&self) -> u16 {
        u16::from_le_bytes([self.bytes[0], self.bytes[1]])
    }

    /// Returns the exact four-byte `op` value.
    #[must_use]
    pub fn raw_value(&self) -> i32 {
        i32::from_le_bytes([self.bytes[2], self.bytes[3], self.bytes[4], self.bytes[5]])
    }

    /// Returns whether the property is a complex property.
    #[must_use]
    pub fn is_complex(&self) -> bool {
        self.raw_opid() & 0x8000 != 0
    }

    /// Returns whether the property is a BLIP-store reference.
    #[must_use]
    pub fn is_blip(&self) -> bool {
        self.raw_opid() & 0x4000 != 0
    }

    /// Returns the exact descriptor and optional complex payload bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Returns the complex payload, or an empty slice for simple properties.
    #[must_use]
    pub fn data(&self) -> &[u8] {
        &self.bytes[6..]
    }
}

/// One unknown `OfficeArt` record retained from a Word shape container.
///
/// `OfficeArt` is extensible and [MS-ODRAW] allows producers to add records that
/// a particular reader does not understand. The record is stored as its exact
/// wire representation, including the eight-byte record header, so a later
/// layer can inspect or replay it without a lossy decode/re-encode cycle.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UnknownRecord {
    bytes: Box<[u8]>,
}

impl UnknownRecord {
    pub(crate) fn from_record(record: &Record<'_>) -> Self {
        Self {
            bytes: record_bytes(record),
        }
    }

    /// Returns the exact `OfficeArt` record bytes, including its header.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Returns the record kind value exactly as it appeared on the wire.
    #[must_use]
    pub fn raw_kind(&self) -> u16 {
        u16::from_le_bytes([self.bytes[2], self.bytes[3]])
    }

    /// Returns the four-bit record version.
    #[must_use]
    pub fn version(&self) -> u8 {
        (u16::from_le_bytes([self.bytes[0], self.bytes[1]]) & 0x000F) as u8
    }

    /// Returns the twelve-bit record instance.
    #[must_use]
    pub fn instance(&self) -> u16 {
        u16::from_le_bytes([self.bytes[0], self.bytes[1]]) >> 4
    }

    /// Returns the record payload without its header.
    #[must_use]
    pub fn data(&self) -> &[u8] {
        &self.bytes[8..]
    }
}

/// Shape information extracted from a Word document's `OfficeArt` drawing.
///
/// The model is owned because textbox text is resolved from separate Word
/// stories after `OfficeArt` parsing. Shape children and unknown-record payloads
/// are consequently independent of the OLE stream borrowed during decoding.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Shape {
    /// Format-neutral `OfficeArt` shape family.
    pub shape_type: Kind,
    /// `OfficeArt` shape identifier (`spid`).
    pub shape_id: u32,
    /// Text content extracted from the shape's Word textbox story, if any.
    pub text: Option<String>,
    /// Whether this shape is a group shape.
    pub is_group: bool,
    /// Child shapes, in `OfficeArt` source order.
    pub children: Vec<Shape>,
    /// Fill color as `(R, G, B)`, when an explicit `fillColor` is present.
    pub fill_color: Option<(u8, u8, u8)>,
    /// Line color as `(R, G, B)`, when an explicit `lineColor` is present.
    pub line_color: Option<(u8, u8, u8)>,
    /// The raw MSOSPT preset-geometry value ([MS-ODRAW] 2.4.24), when present.
    pub native_shape_type: Option<u16>,
    /// Typed `FSPGR` coordinate-space bounds for group shapes.
    pub group_bounds: Option<Bounds>,
    /// Typed `ChildAnchor` coordinates for a shape nested in a group.
    pub anchor: Option<Anchor>,
    /// Host-owned `ClientAnchor` record, retained without reinterpretation.
    pub client_anchor: Option<ClientAnchor>,
    /// Exact flags from the `OfficeArt` `FSP` shape atom.
    pub flags: Flags,
    /// Unknown records directly owned by this shape's `OfficeArt` container.
    pub unknown_records: Vec<UnknownRecord>,
    /// Unknown property entries from the shape's primary, secondary, and
    /// tertiary `OfficeArt` property tables.
    pub unknown_properties: Vec<UnknownProperty>,
    /// Exact wire representation of this shape's `OfficeArt` container.
    ///
    /// For a regular shape this is an `OfficeArtSpContainer`; for a group it
    /// is the enclosing `OfficeArtSpgrContainer`, including its group header
    /// and child records. Keeping the complete container makes replay safe
    /// even when known host records or their original ordering are not part of
    /// the semantic projection.
    pub(crate) office_art: Box<[u8]>,
    pub(crate) text_link: bool,
}

/// A borrowed group projection over one [`Shape`].
///
/// The group view does not duplicate the owned snapshot. It only exposes the
/// topology-specific operations that are meaningful after checking the
/// containing shape's group flag.
#[derive(Debug, Clone, Copy)]
pub struct Group<'a> {
    shape: &'a Shape,
}

impl<'a> Group<'a> {
    /// Returns the group's stable `OfficeArt` identity.
    #[must_use]
    pub const fn identity(self) -> ShapeId {
        ShapeId::from_raw(self.shape.shape_id)
    }

    /// Returns the group's `spid` value.
    #[must_use]
    pub const fn shape_id(self) -> u32 {
        self.shape.shape_id
    }

    /// Returns the group's child-coordinate space.
    #[must_use]
    pub const fn bounds(self) -> Option<&'a Bounds> {
        self.shape.group_bounds.as_ref()
    }

    /// Returns direct children in `OfficeArt` source order.
    #[must_use]
    pub fn children(self) -> &'a [Shape] {
        &self.shape.children
    }

    /// Returns the exact `OfficeArt` group container snapshot, including its
    /// group header and complete child subtree.
    #[must_use]
    pub fn office_art_bytes(self) -> &'a [u8] {
        self.shape.office_art_bytes()
    }

    /// Finds the first descendant with the requested identity in source order.
    #[must_use]
    pub fn find(self, identity: ShapeId) -> Option<&'a Shape> {
        self.shape.find(identity)
    }
}

impl Shape {
    /// Project a host-neutral `OfficeArt` shape into the Word drawing facade.
    pub(crate) fn from_office_art(shape: &OfficeArtShape<'_>) -> io::Result<Self> {
        // In Word, OfficeArtClientTextbox contains only a TXID into the Word
        // textbox story. Text itself is resolved by `Document::text_boxes` and
        // cannot be decoded from OfficeArt bytes in isolation.
        let text_link = if let Some(textbox) = shape.textbox() {
            let _: &[u8; 4] = textbox
                .data()
                .try_into()
                .map_err(|_| invalid_data("Word OfficeArtClientTextbox payload is not one TXID"))?;
            true
        } else {
            false
        };

        let children = shape
            .children()
            .iter()
            .map(Self::from_office_art)
            .collect::<io::Result<_>>()?;

        Ok(Self {
            shape_type: shape.kind(),
            shape_id: shape.id(),
            text: None,
            is_group: matches!(shape.kind(), Kind::Group | Kind::Table),
            children,
            fill_color: shape.props().get_fill_color(),
            line_color: shape.props().get_line_color(),
            native_shape_type: Some(shape.native_kind().raw()),
            group_bounds: shape.group_bounds().copied(),
            anchor: shape.anchor().copied().map(Anchor::from),
            client_anchor: shape.client_anchor().map(ClientAnchor::from_record),
            flags: shape.flags(),
            unknown_records: collect_unknown_records(shape)?,
            unknown_properties: collect_unknown_properties(shape)?,
            office_art: record_bytes(shape.container().record()),
            text_link,
        })
    }

    /// Returns the shape's stable `OfficeArt` identity.
    #[inline]
    #[must_use]
    pub const fn identity(&self) -> ShapeId {
        ShapeId::from_raw(self.shape_id)
    }

    /// Returns the exact `spid` value.
    #[inline]
    #[must_use]
    pub const fn id(&self) -> u32 {
        self.shape_id
    }

    /// Returns the typed `OfficeArt` flags from `FSP`.
    #[inline]
    #[must_use]
    pub const fn shape_flags(&self) -> Flags {
        self.flags
    }

    /// Returns whether the shape is a group according to its `OfficeArt` flag.
    #[inline]
    #[must_use]
    pub const fn is_group_shape(&self) -> bool {
        self.is_group
    }

    /// Returns the typed child anchor, if the shape is nested in a group.
    #[inline]
    #[must_use]
    pub const fn anchor(&self) -> Option<&Anchor> {
        self.anchor.as_ref()
    }

    /// Returns the host-owned client anchor without interpreting its payload.
    #[inline]
    #[must_use]
    pub const fn client_anchor(&self) -> Option<&ClientAnchor> {
        self.client_anchor.as_ref()
    }

    /// Returns the exact `OfficeArt` container used to project this shape.
    ///
    /// The returned bytes are an owned snapshot and can be replayed without
    /// reconstructing the container from the lossy semantic fields. Group
    /// snapshots include their complete child subtree.
    #[must_use]
    pub fn office_art_bytes(&self) -> &[u8] {
        &self.office_art
    }

    /// Returns a typed group view when this shape is a group.
    #[inline]
    #[must_use]
    pub fn group(&self) -> Option<Group<'_>> {
        self.is_group.then_some(Group { shape: self })
    }

    /// Returns direct children in `OfficeArt` source order.
    #[inline]
    #[must_use]
    pub fn children(&self) -> &[Shape] {
        &self.children
    }

    /// Finds the first shape with the requested identity in depth-first source
    /// order. A caller that needs to diagnose producer-invalid duplicate IDs
    /// should inspect every node through [`Self::children`].
    #[must_use]
    pub fn find(&self, identity: ShapeId) -> Option<&Shape> {
        if self.identity() == identity {
            return Some(self);
        }
        self.children.iter().find_map(|child| child.find(identity))
    }

    /// Returns unknown property entries in their property-table order.
    #[inline]
    #[must_use]
    pub fn unknown_properties(&self) -> &[UnknownProperty] {
        &self.unknown_properties
    }
}

fn collect_unknown_records(shape: &OfficeArtShape<'_>) -> io::Result<Vec<UnknownRecord>> {
    let mut records = Vec::new();
    collect_unknown_children(shape.meta(), &mut records)?;

    // For a group shape, `container` wraps the group-header shape and its
    // child shapes. Scan its direct unknown siblings as well, but avoid
    // visiting the same SpContainer twice for ordinary shapes.
    let container = shape.container();
    let meta = shape.meta();
    let same_record = container.record().data().as_ptr() == meta.record().data().as_ptr()
        && container.record().data().len() == meta.record().data().len();
    if !same_record {
        collect_unknown_children(container, &mut records)?;
    }
    Ok(records)
}

fn collect_unknown_properties(shape: &OfficeArtShape<'_>) -> io::Result<Vec<UnknownProperty>> {
    let mut properties = Vec::new();
    for child in shape.meta().children() {
        let child = child.map_err(invalid_data)?;
        let table = match child.kind() {
            RecordKind::Opt => PropertyTable::Primary,
            RecordKind::SecondaryOpt => PropertyTable::Secondary,
            RecordKind::TertiaryOpt => PropertyTable::Tertiary,
            _ => continue,
        };
        let props = Props::parse(&child).map_err(invalid_data)?;
        properties.extend(
            props
                .iter()
                .filter_map(|prop| UnknownProperty::from_prop(table, prop)),
        );
    }
    Ok(properties)
}

fn record_bytes(record: &Record<'_>) -> Box<[u8]> {
    let version_instance = u16::from(record.version()) | (record.instance() << 4);
    let mut bytes = Vec::with_capacity(8usize.saturating_add(record.data().len()));
    bytes.extend_from_slice(&version_instance.to_le_bytes());
    bytes.extend_from_slice(&record.raw_kind().to_le_bytes());
    bytes.extend_from_slice(&record.len().to_le_bytes());
    bytes.extend_from_slice(record.data());
    bytes.into_boxed_slice()
}

fn collect_unknown_children(
    container: &Container<'_>,
    records: &mut Vec<UnknownRecord>,
) -> io::Result<()> {
    for child in container.children() {
        let child = child.map_err(invalid_data)?;
        if matches!(child.kind(), RecordKind::Unknown(_)) {
            records.push(UnknownRecord::from_record(&child));
        }
    }
    Ok(())
}

fn invalid_data(error: impl ToString) -> io::Error {
    io::Error::new(io::ErrorKind::InvalidData, error.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_odraw::Record;

    #[test]
    fn bounds_keep_signed_wire_coordinates() {
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&0x0001_u16.to_le_bytes());
        bytes.extend_from_slice(&0xF009_u16.to_le_bytes());
        bytes.extend_from_slice(&16_u32.to_le_bytes());
        for coordinate in [-120_i32, 40, 960, 880] {
            bytes.extend_from_slice(&coordinate.to_le_bytes());
        }
        let (record, consumed) = Record::parse(&bytes, 0).expect("valid FSPGR record");
        assert_eq!(consumed, bytes.len());
        let bounds = Bounds::from_record(&record).expect("typed FSPGR bounds");
        assert_eq!(bounds, Bounds::new(-120, 40, 960, 880));
        assert_eq!(bounds.width(), Some(1080));
        assert_eq!(bounds.height(), Some(840));
    }

    #[test]
    fn unknown_record_round_trips_exact_wire_bytes() {
        let bytes = [
            0x37, 0x12, // version 7, instance 0x123
            0x34, 0xF1, // producer-defined record kind
            0x03, 0x00, 0x00, 0x00, // payload length
            0xA5, 0x00, 0xFE,
        ];
        let (record, consumed) = Record::parse(&bytes, 0).expect("valid unknown record");
        assert_eq!(consumed, bytes.len());
        let unknown = UnknownRecord::from_record(&record);
        assert_eq!(unknown.bytes(), bytes);
        assert_eq!(unknown.raw_kind(), 0xF134);
        assert_eq!(unknown.version(), 7);
        assert_eq!(unknown.instance(), 0x123);
        assert_eq!(unknown.data(), &bytes[8..]);
    }
}
