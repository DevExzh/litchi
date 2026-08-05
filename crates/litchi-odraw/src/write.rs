//! Checked OfficeArt record writing primitives and builders.

use std::{
    borrow::Cow,
    io::{self, Write},
};

use zerocopy::{
    IntoBytes as _,
    byteorder::little_endian::{U16 as LeU16, U32 as LeU32},
};
use zerocopy_derive::{FromBytes, Immutable, IntoBytes, KnownLayout};

use crate::{
    Limits, Record, RecordKind,
    prop::Id,
    shape::{Flags, Native},
};

/// Property flag indicating that a simple value is a BLIP-store identifier.
pub const BLIP_ID: u16 = 0x4000;
/// Property flag indicating that bytes follow the fixed property table.
pub const COMPLEX: u16 = 0x8000;

/// Frequently used OfficeArt property values.
pub mod prop_value {
    /// Scheme-color marker.
    pub const SCHEME_COLOR: u32 = 0x0800_0000;
    /// Current fill scheme color.
    pub const SCHEME_FILL: u32 = SCHEME_COLOR | 0x04;
    /// Fill-background scheme color.
    pub const SCHEME_FILL_BACK: u32 = SCHEME_COLOR;
    /// Current line scheme color.
    pub const SCHEME_LINE: u32 = SCHEME_COLOR | 0x01;
    /// Current shadow scheme color.
    pub const SCHEME_SHADOW: u32 = SCHEME_COLOR | 0x02;
    /// Default line style.
    pub const LINE_STYLE_DEFAULT: u32 = 0x0010_0010;
    /// Default packed line booleans.
    pub const LINE_STYLE_BOOL_DEFAULT: u32 = 0x0008_0008;
    /// Packed fill-disabled booleans.
    pub const FILL_STYLE_DISABLED: u32 = 0x0010_0000;
    /// Packed fill-enabled booleans.
    pub const FILL_STYLE_ENABLED: u32 = 0x0015_0011;
    /// Packed shadow-disabled booleans.
    pub const SHADOW_STYLE_DISABLED: u32 = 0x0002_0000;
    /// Packed shadow-enabled booleans.
    pub const SHADOW_STYLE_ENABLED: u32 = 0x0002_0002;
}

/// Canonical OfficeArt record identifiers.
pub mod record_type {
    use crate::RecordKind;

    pub const DGG_CONTAINER: u16 = RecordKind::DggContainer.raw();
    pub const DG_CONTAINER: u16 = RecordKind::DgContainer.raw();
    pub const SPGR_CONTAINER: u16 = RecordKind::SpgrContainer.raw();
    pub const SP_CONTAINER: u16 = RecordKind::SpContainer.raw();
    pub const DGG: u16 = RecordKind::Dgg.raw();
    pub const DG: u16 = RecordKind::Dg.raw();
    pub const SPGR: u16 = RecordKind::Spgr.raw();
    pub const SP: u16 = RecordKind::Sp.raw();
    pub const OPT: u16 = RecordKind::Opt.raw();
    pub const CLIENT_TEXTBOX: u16 = RecordKind::ClientTextbox.raw();
    pub const CHILD_ANCHOR: u16 = RecordKind::ChildAnchor.raw();
    pub const CLIENT_ANCHOR: u16 = RecordKind::ClientAnchor.raw();
    pub const CLIENT_DATA: u16 = RecordKind::ClientData.raw();
    pub const SPLIT_MENU_COLORS: u16 = RecordKind::SplitMenuColors.raw();
}

/// Frequently used native `MSOSPT` values.
pub mod shape_type {
    use crate::shape::Native;

    pub const NOT_PRIMITIVE: u16 = Native::FREEFORM.raw();
    pub const RECTANGLE: u16 = Native::RECTANGLE.raw();
    pub const ROUND_RECTANGLE: u16 = Native::ROUND_RECTANGLE.raw();
    pub const ELLIPSE: u16 = Native::ELLIPSE.raw();
    pub const DIAMOND: u16 = Native::DIAMOND.raw();
    pub const LINE: u16 = Native::LINE.raw();
    pub const TEXT_BOX: u16 = Native::TEXT_BOX.raw();
}

/// A validated OfficeArt extension record type.
///
/// Known record types have dedicated variants on [`Atom`] or [`Container`];
/// this route preserves future record types without permitting callers to
/// bypass the invariants attached to known ones.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(transparent)]
pub struct Ext(u16);

impl Ext {
    /// Validates an extension record type in the OfficeArt-reserved range.
    pub fn new(raw: u16) -> io::Result<Self> {
        if raw < 0xF000 {
            return Err(invalid_input("OfficeArt record type is below 0xF000"));
        }
        if !matches!(RecordKind::from_raw(raw), RecordKind::Unknown(_)) {
            return Err(invalid_input(
                "known OfficeArt record type requires a typed variant",
            ));
        }
        Ok(Self(raw))
    }

    /// Returns the exact extension wire value.
    pub const fn raw(self) -> u16 {
        self.0
    }
}

/// A typed OfficeArt atom kind with its required record version.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Atom {
    /// File drawing-group atom.
    Dgg,
    /// BLIP-store entry atom.
    Bse,
    /// Drawing atom.
    Dg,
    /// Shape-group coordinate atom.
    Spgr,
    /// Shape atom.
    Sp,
    /// Primary shape-options atom.
    Opt,
    /// Host client-textbox atom used by DOC and XLS.
    ClientTextbox,
    /// Child-anchor atom.
    ChildAnchor,
    /// Host client-anchor atom.
    ClientAnchor,
    /// Host client-data atom used by DOC and XLS.
    ClientData,
    /// Connector rule atom.
    ConnectorRule,
    /// Alignment rule atom.
    AlignRule,
    /// Arc rule atom.
    ArcRule,
    /// Host client rule atom.
    ClientRule,
    /// Callout rule atom.
    CalloutRule,
    /// Enhanced Metafile BLIP atom.
    BlipEmf,
    /// Windows Metafile BLIP atom.
    BlipWmf,
    /// Macintosh PICT BLIP atom.
    BlipPict,
    /// JPEG BLIP atom using the original record type.
    BlipJpeg,
    /// PNG BLIP atom.
    BlipPng,
    /// Device-independent bitmap BLIP atom.
    BlipDib,
    /// TIFF BLIP atom.
    BlipTiff,
    /// JPEG BLIP atom using the later record-type value.
    BlipJpeg2,
    /// Most-recently-used color atom.
    ColorMru,
    /// Split-menu colors atom.
    SplitMenuColors,
    /// Secondary shape-options atom.
    SecondaryOpt,
    /// Tertiary shape-options atom.
    TertiaryOpt,
    /// A future atom with an explicitly validated non-container version.
    Unknown {
        /// Validated extension record type.
        kind: Ext,
        /// Extension-defined record version.
        version: u8,
    },
}

impl Atom {
    /// Creates a lossless future atom while rejecting known kinds and version 15.
    pub fn unknown(raw: u16, version: u8) -> io::Result<Self> {
        let atom = Self::Unknown {
            kind: Ext::new(raw)?,
            version,
        };
        atom.checked_version()?;
        Ok(atom)
    }

    /// Returns the exact record-type wire value.
    pub const fn raw(self) -> u16 {
        match self {
            Self::Dgg => 0xF006,
            Self::Bse => 0xF007,
            Self::Dg => 0xF008,
            Self::Spgr => 0xF009,
            Self::Sp => 0xF00A,
            Self::Opt => 0xF00B,
            Self::ClientTextbox => 0xF00D,
            Self::ChildAnchor => 0xF00F,
            Self::ClientAnchor => 0xF010,
            Self::ClientData => 0xF011,
            Self::ConnectorRule => 0xF012,
            Self::AlignRule => 0xF013,
            Self::ArcRule => 0xF014,
            Self::ClientRule => 0xF015,
            Self::CalloutRule => 0xF017,
            Self::BlipEmf => 0xF01A,
            Self::BlipWmf => 0xF01B,
            Self::BlipPict => 0xF01C,
            Self::BlipJpeg => 0xF01D,
            Self::BlipPng => 0xF01E,
            Self::BlipDib => 0xF01F,
            Self::BlipTiff => 0xF029,
            Self::BlipJpeg2 => 0xF02A,
            Self::ColorMru => 0xF11A,
            Self::SplitMenuColors => 0xF11E,
            Self::SecondaryOpt => 0xF121,
            Self::TertiaryOpt => 0xF122,
            Self::Unknown { kind, .. } => kind.raw(),
        }
    }

    /// Returns the required record version.
    pub const fn version(self) -> u8 {
        match self {
            Self::Bse | Self::Sp => 2,
            Self::Spgr | Self::ConnectorRule => 1,
            Self::Opt | Self::SecondaryOpt | Self::TertiaryOpt => 3,
            Self::Unknown { version, .. } => version,
            Self::Dgg
            | Self::Dg
            | Self::ClientTextbox
            | Self::ChildAnchor
            | Self::ClientAnchor
            | Self::ClientData
            | Self::AlignRule
            | Self::ArcRule
            | Self::ClientRule
            | Self::CalloutRule
            | Self::BlipEmf
            | Self::BlipWmf
            | Self::BlipPict
            | Self::BlipJpeg
            | Self::BlipPng
            | Self::BlipDib
            | Self::BlipTiff
            | Self::BlipJpeg2
            | Self::ColorMru
            | Self::SplitMenuColors => 0,
        }
    }

    fn checked_version(self) -> io::Result<u8> {
        let version = self.version();
        if version >= 0x0F {
            return Err(invalid_input("atom version must be between 0 and 14"));
        }
        Ok(version)
    }
}

/// A typed OfficeArt container kind.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Container {
    /// File drawing-group container.
    Dgg,
    /// BLIP-store container.
    BStore,
    /// Drawing container.
    Dg,
    /// Shape-group container.
    Spgr,
    /// Shape container.
    Sp,
    /// Solver container.
    Solver,
    /// Host client-textbox container used by PPT.
    ClientTextbox,
    /// Host client-data container used by PPT.
    ClientData,
    /// A future container record type.
    Unknown(Ext),
}

impl Container {
    /// Creates a lossless future container while rejecting known kinds.
    pub fn unknown(raw: u16) -> io::Result<Self> {
        Ext::new(raw).map(Self::Unknown)
    }

    /// Returns the exact record-type wire value.
    pub const fn raw(self) -> u16 {
        match self {
            Self::Dgg => 0xF000,
            Self::BStore => 0xF001,
            Self::Dg => 0xF002,
            Self::Spgr => 0xF003,
            Self::Sp => 0xF004,
            Self::Solver => 0xF005,
            Self::ClientTextbox => 0xF00D,
            Self::ClientData => 0xF011,
            Self::Unknown(kind) => kind.raw(),
        }
    }
}

/// An endian-stable eight-byte OfficeArt record header.
#[derive(Debug, Clone, Copy, FromBytes, IntoBytes, Immutable, KnownLayout)]
#[repr(C)]
pub struct Header {
    ver_inst: LeU16,
    kind: LeU16,
    len: LeU32,
}

impl Header {
    /// Constructs a raw header while applying the OfficeArt bit-field masks.
    ///
    /// This constructor is intentionally infallible for low-level replay and
    /// fixture builders. Checked typed constructors remain available through
    /// [`Header::atom`] and [`Header::container`].
    pub const fn new(version: u8, instance: u16, raw_kind: u16, len: u32) -> Self {
        let ver_inst = ((version as u16) & 0x000F) | ((instance & 0x0FFF) << 4);
        Self {
            ver_inst: LeU16::new(ver_inst),
            kind: LeU16::new(raw_kind),
            len: LeU32::new(len),
        }
    }

    fn from_parts(version: u8, instance: u16, raw_kind: u16, len: u32) -> io::Result<Self> {
        if version > 0x0F {
            return Err(invalid_input("record version exceeds four bits"));
        }
        if instance > 0x0FFF {
            return Err(invalid_input("record instance exceeds twelve bits"));
        }
        Ok(Self::new(version, instance, raw_kind, len))
    }

    /// Constructs a typed atom header with the kind's required version.
    pub fn atom(instance: u16, kind: Atom, len: u32) -> io::Result<Self> {
        Self::from_parts(kind.checked_version()?, instance, kind.raw(), len)
    }

    /// Constructs a version-15 typed container header.
    pub fn container(instance: u16, kind: Container, len: u32) -> io::Result<Self> {
        Self::from_parts(0x0F, instance, kind.raw(), len)
    }

    /// Returns the four-bit version.
    pub fn version(self) -> u8 {
        (self.ver_inst.get() & 0x0F) as u8
    }

    /// Returns the twelve-bit instance.
    pub fn instance(self) -> u16 {
        self.ver_inst.get() >> 4
    }

    /// Returns the record-kind wire value.
    pub fn kind(self) -> RecordKind {
        RecordKind::from_raw(self.kind.get())
    }

    /// Returns the declared payload length.
    pub fn len(self) -> u32 {
        self.len.get()
    }

    /// Returns whether the declared payload is empty.
    pub fn is_empty(self) -> bool {
        self.len() == 0
    }
}

/// Endian-stable shape-atom payload.
#[derive(Debug, Clone, Copy, FromBytes, IntoBytes, Immutable, KnownLayout)]
#[repr(C)]
pub struct Sp {
    id: LeU32,
    flags: LeU32,
}

impl Sp {
    /// Creates a shape payload from typed flags.
    pub const fn new(id: u32, flags: Flags) -> Self {
        Self {
            id: LeU32::new(id),
            flags: LeU32::new(flags.bits()),
        }
    }

    /// Creates a shape payload using the historical builder vocabulary.
    pub const fn with_flags(id: u32, flags: Flags) -> Self {
        Self::new(id, flags)
    }

    /// Creates a group patriarch shape payload.
    pub const fn group_patriarch(id: u32) -> Self {
        Self::new(id, Flags::GROUP.union(Flags::PATRIARCH))
    }

    /// Creates a background shape payload.
    pub const fn background(id: u32) -> Self {
        Self::new(id, Flags::BACKGROUND.union(Flags::HAVE_SPT))
    }

    /// Returns the shape identifier.
    pub fn id(self) -> u32 {
        self.id.get()
    }

    /// Returns all shape flag bits, retaining producer extensions.
    pub fn flags(self) -> Flags {
        Flags::from_bits_retain(self.flags.get())
    }
}

/// Endian-stable six-byte property-table entry.
#[derive(Debug, Clone, Copy, FromBytes, IntoBytes, Immutable, KnownLayout)]
#[repr(C, packed)]
pub struct Property {
    id: LeU16,
    value: LeU32,
}

impl Property {
    /// Creates a raw property-table entry, retaining all flag bits.
    pub const fn new(raw_id: u16, value: u32) -> Self {
        Self {
            id: LeU16::new(raw_id),
            value: LeU32::new(value),
        }
    }

    /// Creates a property-table entry from a typed, unflagged identifier.
    pub const fn from_id(id: Id, value: u32) -> Self {
        Self::new(id.raw(), value)
    }

    /// Returns the exact property identifier, including `fBid`/`fComplex`.
    pub const fn raw_id(self) -> u16 {
        self.id.get()
    }

    /// Returns the typed property identifier without wire flag bits.
    pub fn id(self) -> Id {
        Id::from(self.id.get())
    }

    /// Returns the raw property value.
    pub fn value(self) -> u32 {
        self.value.get()
    }
}

/// Writes one typed atom header without its payload.
pub fn atom_header<W: Write>(
    writer: &mut W,
    instance: u16,
    kind: Atom,
    len: u32,
) -> io::Result<()> {
    writer.write_all(Header::atom(instance, kind, len)?.as_bytes())
}

/// Writes a raw OfficeArt record header for producer-specific replay helpers.
pub fn raw_header<W: Write>(
    writer: &mut W,
    version: u8,
    instance: u16,
    raw_kind: u16,
    len: u32,
) -> io::Result<()> {
    writer.write_all(Header::new(version, instance, raw_kind, len).as_bytes())
}

/// Writes a raw version-15 container without interpreting its child records.
pub fn raw_container<W: Write>(
    writer: &mut W,
    instance: u16,
    raw_kind: u16,
    children: &[u8],
) -> io::Result<()> {
    raw_header(
        writer,
        0x0F,
        instance,
        raw_kind,
        wire_len(children.len(), "container payload")?,
    )?;
    writer.write_all(children)
}

/// Writes a raw OfficeArt atom without interpreting its record kind.
pub fn raw_atom<W: Write>(
    writer: &mut W,
    version: u8,
    instance: u16,
    raw_kind: u16,
    data: &[u8],
) -> io::Result<()> {
    raw_header(
        writer,
        version,
        instance,
        raw_kind,
        wire_len(data.len(), "atom payload")?,
    )?;
    writer.write_all(data)
}

/// Writes one typed container header without its child sequence.
///
/// This is intended for streaming formats such as BIFF that split one
/// container across multiple host records. Prefer [`container`] when the full
/// child sequence is already available so it can be validated before output.
pub fn container_header<W: Write>(
    writer: &mut W,
    instance: u16,
    kind: Container,
    len: u32,
) -> io::Result<()> {
    writer.write_all(Header::container(instance, kind, len)?.as_bytes())
}

/// Writes a container after validating its complete child-record sequence.
pub fn container<W: Write>(
    writer: &mut W,
    instance: u16,
    kind: Container,
    children: &[u8],
) -> io::Result<()> {
    validate_children(children)?;
    let len = wire_len(children.len(), "container payload")?;
    writer.write_all(Header::container(instance, kind, len)?.as_bytes())?;
    writer.write_all(children)
}

/// Writes an atom with the version required by its typed kind.
pub fn atom<W: Write>(writer: &mut W, instance: u16, kind: Atom, data: &[u8]) -> io::Result<()> {
    atom_header(
        writer,
        instance,
        kind,
        wire_len(data.len(), "atom payload")?,
    )?;
    writer.write_all(data)
}

fn validate_children(children: &[u8]) -> io::Result<()> {
    let limits = Limits::default();
    let mut records = 0_u32;
    validate_children_at(children, 0, limits, &mut records)
}

fn validate_children_at(
    children: &[u8],
    depth: u16,
    limits: Limits,
    records: &mut u32,
) -> io::Result<()> {
    if depth > limits.max_depth {
        return Err(invalid_input("child-record nesting exceeds the safe limit"));
    }
    let mut offset = 0_usize;
    while offset < children.len() {
        *records = records
            .checked_add(1)
            .ok_or_else(|| invalid_input("child-record count overflow"))?;
        if *records > limits.max_records {
            return Err(invalid_input("child-record count exceeds the safe limit"));
        }
        let (record, consumed) = Record::parse(children, offset)
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error))?;
        validate_record(&record, depth, limits, records)?;
        offset = offset
            .checked_add(consumed)
            .ok_or_else(|| invalid_input("child-record offset overflow"))?;
    }
    Ok(())
}

fn validate_record(
    record: &Record<'_>,
    depth: u16,
    limits: Limits,
    records: &mut u32,
) -> io::Result<()> {
    let version = record.version();
    let valid = match record.kind() {
        RecordKind::DggContainer
        | RecordKind::BStoreContainer
        | RecordKind::DgContainer
        | RecordKind::SpgrContainer
        | RecordKind::SpContainer
        | RecordKind::SolverContainer => version == 0x0F,
        RecordKind::Dgg
        | RecordKind::Dg
        | RecordKind::ChildAnchor
        | RecordKind::ClientAnchor
        | RecordKind::AlignRule
        | RecordKind::ArcRule
        | RecordKind::ClientRule
        | RecordKind::CalloutRule
        | RecordKind::BlipEmf
        | RecordKind::BlipWmf
        | RecordKind::BlipPict
        | RecordKind::BlipJpeg
        | RecordKind::BlipPng
        | RecordKind::BlipDib
        | RecordKind::BlipTiff
        | RecordKind::ColorMru
        | RecordKind::SplitMenuColors => version == 0,
        RecordKind::Spgr | RecordKind::ConnectorRule => version == 1,
        RecordKind::Bse | RecordKind::Sp => version == 2,
        RecordKind::Opt | RecordKind::SecondaryOpt | RecordKind::TertiaryOpt => version == 3,
        RecordKind::ClientTextbox | RecordKind::ClientData => version == 0 || version == 0x0F,
        RecordKind::Unknown(_) => true,
    };
    if !valid {
        return Err(invalid_input(
            "child record has an invalid version for its known kind",
        ));
    }
    if version == 0x0F {
        let child_depth = depth
            .checked_add(1)
            .ok_or_else(|| invalid_input("child-record nesting depth overflow"))?;
        validate_children_at(record.data(), child_depth, limits, records)?;
    }
    Ok(())
}

#[derive(Debug)]
enum BuiltValue<'data> {
    Simple(i32),
    Blip(i32),
    Complex(Cow<'data, [u8]>),
}

/// Input accepted by the raw property-table builder.
pub trait PropertyKey {
    /// Returns the exact property identifier, including any wire flags.
    fn raw(self) -> u16;
}

impl PropertyKey for Id {
    fn raw(self) -> u16 {
        Id::raw(self)
    }
}

impl PropertyKey for u16 {
    fn raw(self) -> u16 {
        self
    }
}

/// Move-or-borrow builder for an OfficeArt Opt property table.
///
/// Owned complex values are moved into the builder and borrowed values retain
/// their input lifetime. The builder deliberately does not implement `Clone`,
/// preventing an accidental deep copy of owned complex property bytes.
#[derive(Debug, Default)]
pub struct PropertyBuilder<'data> {
    properties: Vec<(u16, BuiltValue<'data>)>,
}

impl<'data> PropertyBuilder<'data> {
    /// Creates an empty property table.
    pub const fn new() -> Self {
        Self {
            properties: Vec::new(),
        }
    }

    /// Appends a simple property.
    pub fn add_simple<K: PropertyKey>(&mut self, id: K, value: i32) -> &mut Self {
        self.properties.push((id.raw(), BuiltValue::Simple(value)));
        self
    }

    /// Appends a BLIP-store identifier property.
    pub fn add_blip_id<K: PropertyKey>(&mut self, id: K, value: i32) -> &mut Self {
        self.properties.push((id.raw(), BuiltValue::Blip(value)));
        self
    }

    /// Appends a complex property by borrowing a slice or moving an owned vector.
    pub fn add_complex<K, D>(&mut self, id: K, data: D) -> &mut Self
    where
        K: PropertyKey,
        D: Into<Cow<'data, [u8]>>,
    {
        self.properties
            .push((id.raw(), BuiltValue::Complex(data.into())));
        self
    }

    /// Returns the exact encoded record size after checked arithmetic.
    pub fn size(&self) -> io::Result<usize> {
        let headers = self
            .properties
            .len()
            .checked_mul(6)
            .ok_or_else(|| invalid_input("property-header size overflow"))?;
        let base = 8_usize
            .checked_add(headers)
            .ok_or_else(|| invalid_input("property record size overflow"))?;
        self.properties
            .iter()
            .try_fold(base, |size, (_, value)| match value {
                BuiltValue::Complex(data) => {
                    i32::try_from(data.len())
                        .map_err(|_| invalid_input("complex property exceeds i32::MAX"))?;
                    size.checked_add(data.len())
                        .ok_or_else(|| invalid_input("property data size overflow"))
                },
                BuiltValue::Simple(_) | BuiltValue::Blip(_) => Ok(size),
            })
    }

    /// Encodes the complete Opt record.
    pub fn write<W: Write>(&self, writer: &mut W) -> io::Result<()> {
        let count = u16::try_from(self.properties.len())
            .map_err(|_| invalid_input("too many OfficeArt properties"))?;
        if count > 0x0FFF {
            return Err(invalid_input("property count exceeds twelve bits"));
        }
        let total = self.size()?;
        let body_len = total
            .checked_sub(8)
            .ok_or_else(|| invalid_input("property record size underflow"))?;
        atom_header(
            writer,
            count,
            Atom::Opt,
            wire_len(body_len, "property payload")?,
        )?;

        for (id, value) in &self.properties {
            let (raw_id, raw_value) = match value {
                BuiltValue::Simple(value) => (*id, *value),
                BuiltValue::Blip(value) => (*id | BLIP_ID, *value),
                BuiltValue::Complex(data) => {
                    let len = i32::try_from(data.len())
                        .map_err(|_| invalid_input("complex property exceeds i32::MAX"))?;
                    (*id | COMPLEX, len)
                },
            };
            writer.write_all(&raw_id.to_le_bytes())?;
            writer.write_all(&raw_value.to_le_bytes())?;
        }
        for (_, value) in &self.properties {
            if let BuiltValue::Complex(data) = value {
                writer.write_all(data)?;
            }
        }
        Ok(())
    }
}

/// Builder for one OfficeArt shape atom.
#[derive(Debug, Clone, Copy)]
pub struct ShapeBuilder {
    kind: Native,
    id: u32,
    flags: Flags,
}

impl ShapeBuilder {
    /// Creates a shape builder with no flags.
    pub fn new<K: Into<Native>>(kind: K, id: u32) -> Self {
        Self {
            kind: kind.into(),
            id,
            flags: Flags::empty(),
        }
    }

    /// Sets typed shape flags.
    pub fn with_flags<F: Into<Flags>>(mut self, flags: F) -> Self {
        self.flags = flags.into();
        self
    }

    /// Encodes the shape atom.
    pub fn write<W: Write>(&self, writer: &mut W) -> io::Result<()> {
        atom_header(writer, self.kind.raw(), Atom::Sp, 8)?;
        writer.write_all(Sp::new(self.id, self.flags).as_bytes())
    }
}

/// Writes a four-coordinate child anchor.
pub fn child_anchor<W: Write>(
    writer: &mut W,
    left: i32,
    top: i32,
    right: i32,
    bottom: i32,
) -> io::Result<()> {
    atom_header(writer, 0, Atom::ChildAnchor, 16)?;
    coordinates(writer, left, top, right, bottom)
}

/// Writes a shape-group coordinate-space atom.
pub fn spgr<W: Write>(
    writer: &mut W,
    left: i32,
    top: i32,
    right: i32,
    bottom: i32,
) -> io::Result<()> {
    atom_header(writer, 0, Atom::Spgr, 16)?;
    coordinates(writer, left, top, right, bottom)
}

/// Writes a drawing atom.
pub fn dg<W: Write>(writer: &mut W, shapes: u32, last_shape_id: u32) -> io::Result<()> {
    atom_header(writer, 0, Atom::Dg, 8)?;
    writer.write_all(&shapes.to_le_bytes())?;
    writer.write_all(&last_shape_id.to_le_bytes())
}

fn coordinates<W: Write>(
    writer: &mut W,
    left: i32,
    top: i32,
    right: i32,
    bottom: i32,
) -> io::Result<()> {
    for coordinate in [left, top, right, bottom] {
        writer.write_all(&coordinate.to_le_bytes())?;
    }
    Ok(())
}

fn wire_len(len: usize, context: &'static str) -> io::Result<u32> {
    u32::try_from(len).map_err(|_| invalid_input(context))
}

fn invalid_input(message: &'static str) -> io::Error {
    io::Error::new(io::ErrorKind::InvalidInput, message)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{Record, RecordKind};
    use std::mem::size_of;

    #[test]
    fn header_is_little_endian_and_checked() {
        let header = Header::atom(0x123, Atom::Sp, 8).expect("valid header");
        assert_eq!(header.as_bytes(), &[0x32, 0x12, 0x0A, 0xF0, 8, 0, 0, 0]);
        assert!(Atom::unknown(0xF123, 0x0F).is_err());
        assert!(Header::atom(0x1000, Atom::Sp, 0).is_err());
        assert!(Ext::new(RecordKind::Sp.raw()).is_err());
        assert!(Ext::new(0x1234).is_err());
        assert_eq!(size_of::<Property>(), 6);
    }

    #[test]
    fn property_builder_moves_or_borrows_complex_data_and_round_trips() {
        let borrowed = [9, 8, 7];
        let moved = vec![1, 2, 3, 4];
        let moved_pointer = moved.as_ptr();
        let mut builder = PropertyBuilder::new();
        builder.add_blip_id(Id::BlipToDisplay, 7);
        builder.add_complex(Id::Vertices, moved);
        builder.add_complex(Id::SegmentInfo, &borrowed[..]);

        match &builder.properties[1].1 {
            BuiltValue::Complex(Cow::Owned(data)) => assert_eq!(data.as_ptr(), moved_pointer),
            value => panic!("expected moved complex bytes, got {value:?}"),
        }
        match &builder.properties[2].1 {
            BuiltValue::Complex(Cow::Borrowed(data)) => {
                assert!(core::ptr::eq(data.as_ptr(), borrowed.as_ptr()));
            },
            value => panic!("expected borrowed complex bytes, got {value:?}"),
        }

        let mut bytes = Vec::new();
        builder.write(&mut bytes).expect("write properties");
        assert_eq!(builder.size().expect("representable size"), 33);
        assert_eq!(u16::from_le_bytes([bytes[8], bytes[9]]), 0x4104);
        assert_eq!(u16::from_le_bytes([bytes[14], bytes[15]]), 0x8145);
        assert_eq!(u16::from_le_bytes([bytes[20], bytes[21]]), 0x8146);
        assert_eq!(&bytes[26..], &[1, 2, 3, 4, 9, 8, 7]);

        let (record, consumed) = Record::parse(&bytes, 0).expect("parse emitted record");
        assert_eq!(record.kind(), RecordKind::Opt);
        assert_eq!(consumed, bytes.len());
    }

    #[test]
    fn container_rejects_malformed_or_mistyped_children_before_writing() {
        let mut output = vec![0xAA];
        let truncated = [0_u8; 7];
        assert!(container(&mut output, 0, Container::Sp, &truncated).is_err());
        assert_eq!(output, [0xAA]);

        let sp_labeled_as_container = [0x0F, 0, 0x0A, 0xF0, 0, 0, 0, 0];
        assert!(container(&mut output, 0, Container::Sp, &sp_labeled_as_container,).is_err());
        assert_eq!(output, [0xAA]);
    }

    #[test]
    fn typed_atom_and_container_round_trip_with_required_versions() {
        let mut child = Vec::new();
        atom(&mut child, 0, Atom::ClientData, &[]).expect("write atom");
        let mut bytes = Vec::new();
        container(&mut bytes, 0, Container::Sp, &child).expect("write container");

        let (root, consumed) = Record::parse(&bytes, 0).expect("parse container");
        assert_eq!(root.kind(), RecordKind::SpContainer);
        assert_eq!(root.version(), 0x0F);
        assert_eq!(consumed, bytes.len());
        let (child, _) = Record::parse(root.data(), 0).expect("parse child");
        assert_eq!(child.kind(), RecordKind::ClientData);
        assert_eq!(child.version(), 0);
    }

    #[test]
    fn extension_kinds_round_trip_losslessly() {
        let mut atom_bytes = Vec::new();
        atom(
            &mut atom_bytes,
            7,
            Atom::unknown(0xF234, 4).expect("valid extension atom"),
            &[1, 2],
        )
        .expect("write extension atom");
        let (atom_record, _) = Record::parse(&atom_bytes, 0).expect("parse extension atom");
        assert_eq!(atom_record.raw_kind(), 0xF234);
        assert_eq!(atom_record.version(), 4);

        let mut container_bytes = Vec::new();
        container(
            &mut container_bytes,
            0,
            Container::unknown(0xF235).expect("valid extension container"),
            &atom_bytes,
        )
        .expect("write extension container");
        let (container_record, _) =
            Record::parse(&container_bytes, 0).expect("parse extension container");
        assert_eq!(container_record.raw_kind(), 0xF235);
        assert_eq!(container_record.version(), 0x0F);
    }

    #[test]
    fn shape_builder_accepts_only_typed_flags() {
        let mut bytes = Vec::new();
        ShapeBuilder::new(Native::RECTANGLE, 42)
            .with_flags(Flags::HAVE_ANCHOR | Flags::HAVE_SPT)
            .write(&mut bytes)
            .expect("write shape");

        let (record, _) = Record::parse(&bytes, 0).expect("parse shape");
        assert_eq!(record.kind(), RecordKind::Sp);
        assert_eq!(record.instance(), Native::RECTANGLE.raw());
    }
}
