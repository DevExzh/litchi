//! Typed, inert Word OLE-control and ObjectPool metadata.

use crate::package::{Error as PackageError, Result};

/// The Word storage containing embedded OLE objects.
pub const OBJECT_POOL_STORAGE: &str = "ObjectPool";
/// The stream containing one object's [`Metadata`] (`ODT`).
pub const OBJ_INFO_STREAM: &str = "\u{3}ObjInfo";
/// The optional stream used by streamed ActiveX controls.
pub const OCX_DATA_STREAM: &str = "\u{3}OCXDATA";
/// The optional screen/print presentation stream for an ObjectPool entry.
pub const PRINT_STREAM: &str = "\u{3}PRINT";
/// The optional Enhanced Metafile print-presentation stream.
pub const EPRINT_STREAM: &str = "\u{3}EPRINT";

/// The location of an OLE-control field in the document stories (the
/// `OcxInfo.idoc` value from MS-DOC 2.9.161).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u16)]
pub enum Story {
    /// Main document story (`idoc = 1`).
    Main = 1,
    /// Header story (`idoc = 2`).
    Header = 2,
    /// Footnote story (`idoc = 3`).
    Footnote = 3,
    /// Textbox story (`idoc = 4`).
    Textbox = 4,
    /// Endnote story (`idoc = 6`).
    Endnote = 6,
    /// Comment story (`idoc = 7`).
    Comment = 7,
    /// Header-textbox story (`idoc = 8`).
    HeaderTextbox = 8,
}

impl Story {
    /// Decode the on-disk `idoc` value.
    pub(crate) fn from_raw(value: u16) -> Result<Self> {
        match value {
            1 => Ok(Self::Main),
            2 => Ok(Self::Header),
            3 => Ok(Self::Footnote),
            4 => Ok(Self::Textbox),
            6 => Ok(Self::Endnote),
            7 => Ok(Self::Comment),
            8 => Ok(Self::HeaderTextbox),
            _ => Err(corrupted(format!("OcxInfo idoc value {value} is invalid"))),
        }
    }

    /// The exact on-disk `idoc` value.
    pub const fn raw(self) -> u16 {
        self as u16
    }
}

/// Counts of fields in the story-specific `Plcfld` tables.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
pub struct FieldCounts {
    main: u32,
    header: u32,
    footnote: u32,
    textbox: u32,
    endnote: u32,
    comment: u32,
    header_textbox: u32,
}

impl FieldCounts {
    /// Construct counts in the story order used by the MS-DOC FIB.
    #[allow(clippy::too_many_arguments)]
    pub const fn new(
        main: u32,
        header: u32,
        footnote: u32,
        textbox: u32,
        endnote: u32,
        comment: u32,
        header_textbox: u32,
    ) -> Self {
        Self {
            main,
            header,
            footnote,
            textbox,
            endnote,
            comment,
            header_textbox,
        }
    }

    /// Return the field count for one `OcxInfo.idoc` story.
    pub const fn for_story(self, story: Story) -> u32 {
        match story {
            Story::Main => self.main,
            Story::Header => self.header,
            Story::Footnote => self.footnote,
            Story::Textbox => self.textbox,
            Story::Endnote => self.endnote,
            Story::Comment => self.comment,
            Story::HeaderTextbox => self.header_textbox,
        }
    }
}

/// The `OcxInfo` flag word.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Flags {
    raw: u16,
}

impl Flags {
    const FIFLD: u16 = 1 << 0;

    /// Construct the flag word from its individual semantic bits.
    #[allow(clippy::too_many_arguments)]
    pub const fn new(
        eats_return: bool,
        eats_escape: bool,
        default_button: bool,
        cancel_button: bool,
        failed_load: bool,
        right_to_left: bool,
        corrupt: bool,
        reserved_bits: u8,
    ) -> Self {
        let mut raw = Self::FIFLD;
        if eats_return {
            raw |= 1 << 1;
        }
        if eats_escape {
            raw |= 1 << 2;
        }
        if default_button {
            raw |= 1 << 3;
        }
        if cancel_button {
            raw |= 1 << 4;
        }
        if failed_load {
            raw |= 1 << 5;
        }
        if right_to_left {
            raw |= 1 << 6;
        }
        if corrupt {
            raw |= 1 << 7;
        }
        Self {
            raw: raw | (reserved_bits as u16) << 8,
        }
    }

    /// Decode the raw flag word, enforcing the required `fifld` bit.
    pub(crate) fn from_raw(raw: u16) -> Result<Self> {
        super::validation::flags(raw)?;
        Ok(Self { raw })
    }

    /// The exact serialized flag word.
    pub const fn raw(self) -> u16 {
        self.raw
    }

    /// Whether the record is associated with a field (`fifld`).
    pub const fn field_present(self) -> bool {
        self.raw & Self::FIFLD != 0
    }

    /// Whether the control consumes ENTER.
    pub const fn eats_return(self) -> bool {
        self.raw & (1 << 1) != 0
    }

    /// Whether the control consumes ESC.
    pub const fn eats_escape(self) -> bool {
        self.raw & (1 << 2) != 0
    }

    /// Whether the control is the default button.
    pub const fn default_button(self) -> bool {
        self.raw & (1 << 3) != 0
    }

    /// Whether the control is the default CANCEL button.
    pub const fn cancel_button(self) -> bool {
        self.raw & (1 << 4) != 0
    }

    /// Whether loading the control failed.
    pub const fn failed_load(self) -> bool {
        self.raw & (1 << 5) != 0
    }

    /// Whether the control uses right-to-left display handling.
    pub const fn right_to_left(self) -> bool {
        self.raw & (1 << 6) != 0
    }

    /// Whether the control is marked corrupt.
    pub const fn corrupt(self) -> bool {
        self.raw & (1 << 7) != 0
    }

    /// The ignored high-byte bits, retained losslessly.
    pub const fn reserved_bits(self) -> u8 {
        (self.raw >> 8) as u8
    }
}

/// One fixed-size `OcxInfo` record (MS-DOC 2.9.161).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct OcxInfo {
    cookie: u32,
    field_index: u32,
    accelerator_handle: u32,
    accelerator_count: u16,
    flags: Flags,
    story: Story,
    reserved: u16,
}

impl OcxInfo {
    /// Construct a record. The ignored/reserved values are retained exactly
    /// when the record is serialized.
    pub const fn new(
        cookie: u32,
        field_index: u32,
        accelerator_handle: u32,
        accelerator_count: u16,
        flags: Flags,
        story: Story,
        reserved: u16,
    ) -> Self {
        Self {
            cookie,
            field_index,
            accelerator_handle,
            accelerator_count,
            flags,
            story,
            reserved,
        }
    }

    /// Validate the record independently of its containing array.
    pub fn validate(self) -> Result<()> {
        super::validation::info(&self)
    }

    /// Unique `dwCookie` index in the containing table.
    pub const fn cookie(self) -> u32 {
        self.cookie
    }

    /// `ifld`, the field index in the story selected by [`Self::story`].
    pub const fn field_index(self) -> u32 {
        self.field_index
    }

    /// Undefined `hAccel`, retained without interpretation.
    pub const fn accelerator_handle(self) -> u32 {
        self.accelerator_handle
    }

    /// Number of accelerator entries (`cAccel`).
    pub const fn accelerator_count(self) -> u16 {
        self.accelerator_count
    }

    /// Semantic and retained bits from the record flag word.
    pub const fn flags(self) -> Flags {
        self.flags
    }

    /// Story containing the field referenced by `ifld`.
    pub const fn story(self) -> Story {
        self.story
    }

    /// Undefined `reserved2`, retained without interpretation.
    pub const fn reserved(self) -> u16 {
        self.reserved
    }
}

/// The `RgxOcxInfo` array (MS-DOC 2.9.229).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RgxOcxInfo {
    infos: Vec<OcxInfo>,
}

impl RgxOcxInfo {
    /// Construct an array and validate the document-wide cookie invariant.
    pub fn try_new(infos: Vec<OcxInfo>) -> Result<Self> {
        super::validation::infos(&infos)?;
        Ok(Self { infos })
    }

    pub(crate) fn from_infos(infos: Vec<OcxInfo>) -> Self {
        Self { infos }
    }

    /// Records in their original table order.
    pub fn infos(&self) -> &[OcxInfo] {
        &self.infos
    }

    /// Number of records.
    pub fn len(&self) -> usize {
        self.infos.len()
    }

    /// Whether no records are present.
    pub fn is_empty(&self) -> bool {
        self.infos.is_empty()
    }

    pub(crate) fn validate(&self) -> Result<()> {
        super::validation::infos(&self.infos)
    }

    /// Validate each `ifld` against the story-specific `Plcfld` count.
    pub fn validate_fields(&self, counts: FieldCounts) -> Result<()> {
        super::validation::field_indices(&self.infos, counts)
    }
}

/// The format identifier (`cf`) in an ObjectPool `ODT` record.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u16)]
pub enum Format {
    /// Rich Text Format (`0x0001`).
    RichText = 0x0001,
    /// Plain text (`0x0002`).
    Text = 0x0002,
    /// Metafile or Enhanced Metafile (`0x0003`).
    Metafile = 0x0003,
    /// Bitmap (`0x0004`).
    Bitmap = 0x0004,
    /// Device-independent bitmap (`0x0005`).
    DeviceIndependentBitmap = 0x0005,
    /// HTML (`0x000A`).
    Html = 0x000A,
    /// Unicode text (`0x0014`).
    UnicodeText = 0x0014,
}

impl Format {
    pub(crate) fn from_raw(value: u16) -> Result<Self> {
        match value {
            0x0001 => Ok(Self::RichText),
            0x0002 => Ok(Self::Text),
            0x0003 => Ok(Self::Metafile),
            0x0004 => Ok(Self::Bitmap),
            0x0005 => Ok(Self::DeviceIndependentBitmap),
            0x000A => Ok(Self::Html),
            0x0014 => Ok(Self::UnicodeText),
            _ => Err(corrupted(format!("ODT cf value {value:#06x} is invalid"))),
        }
    }

    /// The exact on-disk `cf` value.
    pub const fn raw(self) -> u16 {
        self as u16
    }
}

/// The `ODTPersist1` bitfield in an ObjectPool `ODT` record.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Persist1 {
    raw: u16,
}

impl Persist1 {
    /// Construct a validated bitfield while retaining undefined bits.
    #[allow(clippy::too_many_arguments)]
    pub fn try_new(
        default_handler: bool,
        linked: bool,
        display_as_icon: bool,
        ole1: bool,
        manual_update: bool,
        recompose_on_resize: bool,
        activex: bool,
        stream_data: bool,
        view_object: bool,
        reserved_bits: u16,
    ) -> Result<Self> {
        let mut raw = reserved_bits;
        raw |= bit(default_handler, 1);
        raw |= bit(linked, 4);
        raw |= bit(display_as_icon, 6);
        raw |= bit(ole1, 7);
        raw |= bit(manual_update, 8);
        raw |= bit(recompose_on_resize, 9);
        raw |= bit(activex, 12);
        raw |= bit(stream_data, 13);
        raw |= bit(view_object, 15);
        Self::from_raw(raw)
    }

    pub(crate) fn from_raw(raw: u16) -> Result<Self> {
        super::validation::persist1(raw)?;
        Ok(Self { raw })
    }

    /// The exact serialized bitfield.
    pub const fn raw(self) -> u16 {
        self.raw
    }

    /// Whether Word should assume its default document handler CLSID.
    pub const fn default_handler(self) -> bool {
        self.raw & (1 << 1) != 0
    }

    /// Whether the OLE object is a link.
    pub const fn linked(self) -> bool {
        self.raw & (1 << 4) != 0
    }

    /// Whether the object is represented by an icon.
    pub const fn display_as_icon(self) -> bool {
        self.raw & (1 << 6) != 0
    }

    /// Whether the object is OLE 1-only.
    pub const fn ole1(self) -> bool {
        self.raw & (1 << 7) != 0
    }

    /// Whether a link updates only on user action.
    pub const fn manual_update(self) -> bool {
        self.raw & (1 << 8) != 0
    }

    /// Whether the object requests resize notifications.
    pub const fn recompose_on_resize(self) -> bool {
        self.raw & (1 << 9) != 0
    }

    /// Whether the object is an OLE control.
    pub const fn activex(self) -> bool {
        self.raw & (1 << 12) != 0
    }

    /// Whether an ActiveX control stores its data in `OCXDATA`.
    pub const fn stream_data(self) -> bool {
        self.raw & (1 << 13) != 0
    }

    /// Whether the object supports `IViewObject`.
    pub const fn view_object(self) -> bool {
        self.raw & (1 << 15) != 0
    }

    /// Undefined bits retained from `ODTPersist1`.
    pub const fn reserved_bits(self) -> u16 {
        self.raw & super::validation::PERSIST1_UNDEFINED
    }
}

/// The `ODTPersist2` bitfield in an ObjectPool `ODT` record.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Persist2 {
    raw: u16,
}

impl Persist2 {
    /// Construct a validated bitfield while retaining undefined bits.
    pub fn try_new(
        enhanced_metafile: bool,
        queried_enhanced_metafile: bool,
        stored_as_enhanced_metafile: bool,
        reserved_bits: u16,
    ) -> Result<Self> {
        let mut raw = reserved_bits;
        raw |= bit(enhanced_metafile, 0);
        raw |= bit(queried_enhanced_metafile, 2);
        raw |= bit(stored_as_enhanced_metafile, 3);
        Self::from_raw(raw)
    }

    pub(crate) fn from_raw(raw: u16) -> Result<Self> {
        super::validation::persist2(raw)?;
        Ok(Self { raw })
    }

    /// The exact serialized bitfield.
    pub const fn raw(self) -> u16 {
        self.raw
    }

    /// Whether the document presentation is Enhanced Metafile.
    pub const fn enhanced_metafile(self) -> bool {
        self.raw & (1 << 0) != 0
    }

    /// Whether the producer queried Enhanced Metafile support.
    pub const fn queried_enhanced_metafile(self) -> bool {
        self.raw & (1 << 2) != 0
    }

    /// Whether the object supports Enhanced Metafile data.
    pub const fn stored_as_enhanced_metafile(self) -> bool {
        self.raw & (1 << 3) != 0
    }

    /// Undefined bits retained from `ODTPersist2`.
    pub const fn reserved_bits(self) -> u16 {
        self.raw & super::validation::PERSIST2_UNDEFINED
    }
}

/// Typed `ODT` metadata from one ObjectPool storage.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Metadata {
    persist1: Persist1,
    format: Format,
    persist2: Option<Persist2>,
}

impl Metadata {
    /// Construct and validate one ObjectPool metadata record.
    pub fn try_new(persist1: Persist1, format: Format, persist2: Option<Persist2>) -> Result<Self> {
        let value = Self {
            persist1,
            format,
            persist2,
        };
        value.validate()?;
        Ok(value)
    }

    /// Validate the metadata without opening or interpreting the object.
    pub fn validate(self) -> Result<()> {
        super::validation::metadata(&self)
    }

    /// The required `ODTPersist1` bitfield.
    pub const fn persist1(self) -> Persist1 {
        self.persist1
    }

    /// The `cf` presentation format.
    pub const fn format(self) -> Format {
        self.format
    }

    /// The optional `ODTPersist2` bitfield, preserving presence separately
    /// from an all-zero value.
    pub const fn persist2(self) -> Option<Persist2> {
        self.persist2
    }

    /// Whether this ObjectPool entry identifies an ActiveX/OLE control.
    pub const fn is_activex(self) -> bool {
        self.persist1.activex()
    }

    /// Whether the ActiveX data is expected in `OCXDATA`.
    pub const fn stores_control_data_in_stream(self) -> bool {
        self.persist1.stream_data()
    }
}

/// A validated ObjectPool storage name, such as `_42` or `_-1`.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct StorageName(String);

impl StorageName {
    /// Construct a name using the decimal storage-name grammar from MS-DOC.
    pub fn try_new(name: impl Into<String>) -> Result<Self> {
        let name = name.into();
        super::validation::storage_name(&name)?;
        Ok(Self(name))
    }

    /// Borrow the exact storage name.
    pub fn as_str(&self) -> &str {
        &self.0
    }

    /// Consume the name and return its exact spelling.
    pub fn into_string(self) -> String {
        self.0
    }
}

/// Passive ActiveX state derived from an ObjectPool entry.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ActiveX {
    stream_data: bool,
    data_present: bool,
}

impl ActiveX {
    pub(crate) const fn new(stream_data: bool, data_present: bool) -> Self {
        Self {
            stream_data,
            data_present,
        }
    }

    /// Whether the control stores its data in `OCXDATA`.
    pub const fn stream_data(self) -> bool {
        self.stream_data
    }

    /// Whether the selected ObjectPool storage contains `OCXDATA`.
    pub const fn data_present(self) -> bool {
        self.data_present
    }
}

/// One inert ObjectPool storage entry and its recognized metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Entry {
    name: StorageName,
    class_id: Option<String>,
    metadata: Option<Metadata>,
    control_data_present: bool,
    print_present: bool,
    enhanced_print_present: bool,
}

impl Entry {
    /// Construct and validate one passive ObjectPool entry.
    pub fn try_new(
        name: StorageName,
        class_id: Option<String>,
        metadata: Option<Metadata>,
        control_data_present: bool,
    ) -> Result<Self> {
        Self::try_with_streams(name, class_id, metadata, control_data_present, false, false)
    }

    /// Construct one entry while retaining the optional presentation-stream
    /// presence bits from the selected CFB storage.
    pub fn try_with_streams(
        name: StorageName,
        class_id: Option<String>,
        metadata: Option<Metadata>,
        control_data_present: bool,
        print_present: bool,
        enhanced_print_present: bool,
    ) -> Result<Self> {
        let value = Self {
            name,
            class_id,
            metadata,
            control_data_present,
            print_present,
            enhanced_print_present,
        };
        value.validate()?;
        Ok(value)
    }

    /// Validate the storage-name, `ODT`, and ActiveX stream relationship.
    pub fn validate(&self) -> Result<()> {
        super::validation::entry(self)
    }

    /// The exact ObjectPool storage name.
    pub fn name(&self) -> &StorageName {
        &self.name
    }

    /// The CFB storage CLSID, when present in the captured directory entry.
    pub fn class_id(&self) -> Option<&str> {
        self.class_id.as_deref()
    }

    /// The typed `\x03ObjInfo` metadata, when the stream was present.
    pub const fn metadata(&self) -> Option<&Metadata> {
        self.metadata.as_ref()
    }

    /// Whether the entry is marked as an ActiveX/OLE control.
    pub fn is_activex(&self) -> bool {
        self.metadata.is_some_and(|value| value.is_activex())
    }

    /// Passive ActiveX metadata, without opening or executing its payload.
    pub fn active_x(&self) -> Option<ActiveX> {
        self.metadata
            .filter(|value| value.is_activex())
            .map(|value| {
                ActiveX::new(
                    value.stores_control_data_in_stream(),
                    self.control_data_present,
                )
            })
    }

    /// Whether the `\x03OCXDATA` stream was present.
    pub const fn control_data_present(&self) -> bool {
        self.control_data_present
    }

    /// Whether the `\x03PRINT` presentation stream was present.
    pub const fn print_present(&self) -> bool {
        self.print_present
    }

    /// Whether the `\x03EPRINT` presentation stream was present.
    pub const fn enhanced_print_present(&self) -> bool {
        self.enhanced_print_present
    }
}

/// An immutable, ordered inventory of selected ObjectPool storages.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ObjectPool {
    entries: Vec<Entry>,
}

impl ObjectPool {
    /// Construct and validate an ObjectPool inventory.
    pub fn try_new(entries: Vec<Entry>) -> Result<Self> {
        super::validation::pool(&entries)?;
        Ok(Self { entries })
    }

    /// Entries in their original CFB discovery order.
    pub fn entries(&self) -> &[Entry] {
        &self.entries
    }

    /// Number of selected ObjectPool storages.
    pub const fn len(&self) -> usize {
        self.entries.len()
    }

    /// Whether the inventory is empty.
    pub const fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    /// Find an entry by its exact storage name.
    pub fn get(&self, name: &str) -> Option<&Entry> {
        self.entries
            .iter()
            .find(|entry| entry.name().as_str() == name)
    }
}

fn bit(value: bool, shift: u16) -> u16 {
    if value { 1 << shift } else { 0 }
}

fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}
