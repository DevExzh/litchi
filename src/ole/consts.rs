/// Magic bytes that should be at the beginning of every OLE file
pub const MAGIC: &[u8; 8] = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1";

/// Minimal size of an empty OLE file with 512-byte sectors (1536 bytes)
pub const MINIMAL_OLEFILE_SIZE: usize = 1536;

/// Size of a directory entry in bytes
pub const DIRENTRY_SIZE: usize = 128;

/// Default sector size for version 3 (512 bytes)
pub const SECTOR_SIZE_V3: usize = 512;

/// Default sector size for version 4 (4096 bytes)
pub const SECTOR_SIZE_V4: usize = 4096;

// Sector IDs (from AAF specifications)
/// Maximum regular sector ID
pub const MAXREGSECT: u32 = 0xFFFFFFFA; // -6
/// Denotes a DIFAT sector in a FAT
pub const DIFSECT: u32 = 0xFFFFFFFC; // -4
/// Denotes a FAT sector in a FAT
pub const FATSECT: u32 = 0xFFFFFFFD; // -3
/// End of a virtual stream chain
pub const ENDOFCHAIN: u32 = 0xFFFFFFFE; // -2
/// Unallocated sector
pub const FREESECT: u32 = 0xFFFFFFFF; // -1

// Directory Entry IDs (from AAF specifications)
/// Maximum directory entry ID
pub const MAXREGSID: u32 = 0xFFFFFFFA; // -6
/// Unallocated directory entry
pub const NOSTREAM: u32 = 0xFFFFFFFF; // -1

// Object types in storage (from AAF specifications)
/// Empty directory entry
pub const STGTY_EMPTY: u8 = 0;
/// Element is a storage object
pub const STGTY_STORAGE: u8 = 1;
/// Element is a stream object
pub const STGTY_STREAM: u8 = 2;
/// Element is an ILockBytes object
pub const STGTY_LOCKBYTES: u8 = 3;
/// Element is an IPropertyStorage object
pub const STGTY_PROPERTY: u8 = 4;
/// Element is a root storage
pub const STGTY_ROOT: u8 = 5;

/// Unknown size for a stream (used when size is not known in advance)
pub const UNKNOWN_SIZE: u32 = 0x7FFFFFFF;

// Property types
pub const VT_EMPTY: u16 = 0;
pub const VT_NULL: u16 = 1;
pub const VT_I2: u16 = 2;
pub const VT_I4: u16 = 3;
pub const VT_R4: u16 = 4;
pub const VT_R8: u16 = 5;
pub const VT_CY: u16 = 6;
pub const VT_DATE: u16 = 7;
pub const VT_BSTR: u16 = 8;
pub const VT_DISPATCH: u16 = 9;
pub const VT_ERROR: u16 = 10;
pub const VT_BOOL: u16 = 11;
pub const VT_VARIANT: u16 = 12;
pub const VT_UNKNOWN: u16 = 13;
pub const VT_DECIMAL: u16 = 14;
pub const VT_I1: u16 = 16;
pub const VT_UI1: u16 = 17;
pub const VT_UI2: u16 = 18;
pub const VT_UI4: u16 = 19;
pub const VT_I8: u16 = 20;
pub const VT_UI8: u16 = 21;
pub const VT_INT: u16 = 22;
pub const VT_UINT: u16 = 23;
pub const VT_VOID: u16 = 24;
pub const VT_HRESULT: u16 = 25;
pub const VT_PTR: u16 = 26;
pub const VT_SAFEARRAY: u16 = 27;
pub const VT_CARRAY: u16 = 28;
pub const VT_USERDEFINED: u16 = 29;
pub const VT_LPSTR: u16 = 30;
pub const VT_LPWSTR: u16 = 31;
pub const VT_FILETIME: u16 = 64;
pub const VT_BLOB: u16 = 65;
pub const VT_STREAM: u16 = 66;
pub const VT_STORAGE: u16 = 67;
pub const VT_STREAMED_OBJECT: u16 = 68;
pub const VT_STORED_OBJECT: u16 = 69;
pub const VT_BLOB_OBJECT: u16 = 70;
pub const VT_CF: u16 = 71;
pub const VT_CLSID: u16 = 72;
pub const VT_VECTOR: u16 = 0x1000;

/// Common document type: Microsoft Word
pub const WORD_CLSID: &str = "00020900-0000-0000-C000-000000000046";

// PowerPoint Binary File Format (MS-PPT) constants

/// PPT record types (based on POI RecordTypes enum)
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u16)]
pub enum PptRecordType {
    /// Unknown record type
    Unknown = 0,
    /// Document record
    Document = 1000,
    /// Document atom record
    DocumentAtom = 1001,
    /// End document record
    EndDocument = 1002,
    /// Slide record
    Slide = 1006,
    /// Slide atom record
    SlideAtom = 1007,
    /// Notes record
    Notes = 1008,
    /// Notes atom record
    NotesAtom = 1009,
    /// Environment record
    Environment = 1010,
    /// Slide persist atom record
    SlidePersistAtom = 1011,
    /// Main master record
    MainMaster = 1016,
    /// Slide list with text record
    SlideListWithText = 4080,
    /// Persist pointer holder record
    PersistPtrHolder = 6001,
    /// Slide show slide info atom
    SSSlideInfoAtom = 1017,
    /// VBA info record
    VBAInfo = 1023,
    /// VBA info atom record
    VBAInfoAtom = 1024,
    /// External object list record
    ExObjList = 1033,
    /// External object list atom record
    ExObjListAtom = 1034,
    /// PP drawing group record
    PPDrawingGroup = 1035,
    /// PP drawing record
    PPDrawing = 1036,
    /// OE placeholder atom record (placeholder data)
    OEPlaceholderAtom = 3011,
    /// Text header atom record
    TextHeaderAtom = 3999,
    /// Text characters atom record
    TextCharsAtom = 4000,
    /// Text bytes atom record
    TextBytesAtom = 4008,
    /// Text special info atom record
    TextSpecInfoAtom = 4010,
    /// Style text prop atom record
    StyleTextPropAtom = 4001,
    /// Master text prop atom record
    MasterTextPropAtom = 4002,
    /// Text master style atom record
    TxMasterStyleAtom = 4003,
    /// Text CF style atom record
    TxCFStyleAtom = 4004,
    /// Text PF style atom record
    TxPFStyleAtom = 4005,
    /// Text ruler atom record
    TextRulerAtom = 4006,
    /// Font entity atom record
    FontEntityAtom = 4023,
    /// CString record
    CString = 4026,
    /// Headers footers container record
    HeadersFooters = 4057,
    /// Headers footers atom record
    HeadersFootersAtom = 4058,
    /// Interactive info record
    InteractiveInfo = 4082,
    /// Interactive info atom record
    InteractiveInfoAtom = 4083,
    /// User edit atom record
    UserEditAtom = 4085,
    /// Current user atom record
    CurrentUserAtom = 4086,
    /// Date time MC atom record
    DateTimeMCAtom = 4087,
    /// Animation info record
    AnimationInfo = 4116,
    /// Animation info atom record
    AnimationInfoAtom = 4081,
    /// Build list record
    BuildList = 2000,
    /// Build atom record
    BuildAtom = 2001,
    /// Chart build record
    ChartBuild = 2010,
    /// Diagram build record
    DiagramBuild = 2011,
    /// Paragraph build record
    ParaBuild = 2012,
    /// Sound collection container
    SoundCollection = 2020,
    /// Sound collection atom
    SoundCollectionAtom = 2021,
    /// Sound record
    Sound = 2022,
    /// Sound data record
    SoundData = 2023,
    /// Time node record
    TimeNode = 4114,
    /// Time property list record
    TimePropertyList = 4115,
    /// Time behavior record
    TimeBehavior = 4112,
    /// Comment 2000 record
    Comment2000 = 12000,
    /// Comment 2000 atom record
    Comment2000Atom = 12001,
}

impl From<u16> for PptRecordType {
    fn from(value: u16) -> Self {
        match value {
            0 => PptRecordType::Unknown,
            1000 => PptRecordType::Document,
            1001 => PptRecordType::DocumentAtom,
            1002 => PptRecordType::EndDocument,
            1006 => PptRecordType::Slide,
            1007 => PptRecordType::SlideAtom,
            1008 => PptRecordType::Notes,
            1009 => PptRecordType::NotesAtom,
            1010 => PptRecordType::Environment,
            1011 => PptRecordType::SlidePersistAtom,
            1016 => PptRecordType::MainMaster,
            1017 => PptRecordType::SSSlideInfoAtom,
            4080 => PptRecordType::SlideListWithText,
            6001 | 6002 => PptRecordType::PersistPtrHolder, // Both values are used
            1023 => PptRecordType::VBAInfo,
            1024 => PptRecordType::VBAInfoAtom,
            1033 => PptRecordType::ExObjList,
            1034 => PptRecordType::ExObjListAtom,
            1035 => PptRecordType::PPDrawingGroup,
            1036 => PptRecordType::PPDrawing,
            3011 => PptRecordType::OEPlaceholderAtom,
            3999 => PptRecordType::TextHeaderAtom,
            4000 => PptRecordType::TextCharsAtom,
            4008 => PptRecordType::TextBytesAtom,
            4010 => PptRecordType::TextSpecInfoAtom,
            4001 => PptRecordType::StyleTextPropAtom,
            4002 => PptRecordType::MasterTextPropAtom,
            4003 => PptRecordType::TxMasterStyleAtom,
            4004 => PptRecordType::TxCFStyleAtom,
            4005 => PptRecordType::TxPFStyleAtom,
            4006 => PptRecordType::TextRulerAtom,
            4023 => PptRecordType::FontEntityAtom,
            4026 => PptRecordType::CString,
            4057 => PptRecordType::HeadersFooters,
            4058 => PptRecordType::HeadersFootersAtom,
            4082 => PptRecordType::InteractiveInfo,
            4083 => PptRecordType::InteractiveInfoAtom,
            4085 => PptRecordType::UserEditAtom,
            4086 => PptRecordType::CurrentUserAtom,
            4087 => PptRecordType::DateTimeMCAtom,
            4116 => PptRecordType::AnimationInfo,
            4081 => PptRecordType::AnimationInfoAtom,
            2000 => PptRecordType::BuildList,
            2001 => PptRecordType::BuildAtom,
            2010 => PptRecordType::ChartBuild,
            2011 => PptRecordType::DiagramBuild,
            2012 => PptRecordType::ParaBuild,
            4114 => PptRecordType::TimeNode,
            4115 => PptRecordType::TimePropertyList,
            4112 => PptRecordType::TimeBehavior,
            12000 => PptRecordType::Comment2000,
            12001 => PptRecordType::Comment2000Atom,
            _ => PptRecordType::Unknown,
        }
    }
}

impl PptRecordType {
    /// Get the u16 value of this record type
    pub fn as_u16(self) -> u16 {
        unsafe { std::mem::transmute::<Self, u16>(self) }
    }
}

/// Escher record types (MS-ODRAW format)
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u16)]
pub enum EscherRecordType {
    /// Container record
    Container = 0xF000,
    /// Shape record
    Shape = 0xF004,
    /// Text box record
    TextBox = 0xF00C,
    /// Client text box record
    ClientTextBox = 0xF00D,
    /// Child anchor record
    ChildAnchor = 0xF00E,
    /// Client anchor record
    ClientAnchor = 0xF00F,
    /// Client data record
    ClientData = 0xF010,
    /// Properties record
    Properties = 0xF011,
    /// Transform record
    Transform = 0xF012,
    /// Text record
    Text = 0xF013,
    /// Placeholder data record
    PlaceholderData = 0xF014,
}

impl From<u16> for EscherRecordType {
    fn from(value: u16) -> Self {
        match value {
            0xF000 => EscherRecordType::Container,
            0xF004 => EscherRecordType::Shape,
            0xF00C => EscherRecordType::TextBox,
            0xF00D => EscherRecordType::ClientTextBox,
            0xF00E => EscherRecordType::ChildAnchor,
            0xF00F => EscherRecordType::ClientAnchor,
            0xF010 => EscherRecordType::ClientData,
            0xF011 => EscherRecordType::Properties,
            0xF012 => EscherRecordType::Transform,
            0xF013 => EscherRecordType::Text,
            0xF014 => EscherRecordType::PlaceholderData,
            _ => EscherRecordType::Container, // Default fallback
        }
    }
}

impl EscherRecordType {
    /// Get the u16 value of this record type
    pub fn as_u16(self) -> u16 {
        unsafe { std::mem::transmute::<Self, u16>(self) }
    }
}

// Additional Escher/MS-ODRAW constants

/// Escher record version bits (high 12 bits)
pub const ESCHER_VERSION_MASK: u16 = 0x0FFF;

/// Escher record instance bits (low 12 bits)
pub const ESCHER_INSTANCE_MASK: u16 = 0x0FFF;

/// Escher record header size in bytes
pub const ESCHER_HEADER_SIZE: usize = 8;

/// Minimum size for a valid Escher record
pub const ESCHER_MIN_RECORD_SIZE: usize = ESCHER_HEADER_SIZE;

/// Escher container record flag (has children)
pub const ESCHER_CONTAINER_FLAG: u16 = 0x000F;

/// Shape types in Escher format
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u16)]
pub enum EscherShapeType {
    /// Not a primitive shape
    NotPrimitive = 0,
    /// Rectangle
    Rectangle = 1,
    /// Round rectangle
    RoundRectangle = 2,
    /// Oval
    Oval = 3,
    /// Diamond
    Diamond = 4,
    /// Isosceles triangle
    Triangle = 5,
    /// Right triangle
    RightTriangle = 6,
    /// Parallelogram
    Parallelogram = 7,
    /// Trapezoid
    Trapezoid = 8,
    /// Hexagon
    Hexagon = 9,
    /// Octagon
    Octagon = 10,
    /// Plus sign
    Plus = 11,
    /// Star
    Star = 12,
    /// Arrow
    Arrow = 13,
    /// Thick arrow
    ThickArrow = 14,
    /// Home plate
    HomePlate = 15,
    /// Cube
    Cube = 16,
    /// Balloon
    Balloon = 17,
    /// Seal
    Seal = 18,
    /// Arc
    Arc = 19,
    /// Line
    Line = 20,
    /// Plaque
    Plaque = 21,
    /// Can
    Can = 22,
    /// Donut
    Donut = 23,
    /// Text simple
    TextSimple = 24,
    /// Text octagon
    TextOctagon = 25,
    /// Text hexagon
    TextHexagon = 26,
    /// Text curve
    TextCurve = 27,
    /// Text wave
    TextWave = 28,
    /// Text ring
    TextRing = 29,
    /// Text on curve
    TextOnCurve = 30,
    /// Text on ring
    TextOnRing = 31,
    /// Straight connector 1
    StraightConnector1 = 32,
    /// Bent connector 2
    BentConnector2 = 33,
    /// Bent connector 3
    BentConnector3 = 34,
    /// Bent connector 4
    BentConnector4 = 35,
    /// Bent connector 5
    BentConnector5 = 36,
    /// Curved connector 2
    CurvedConnector2 = 37,
    /// Curved connector 3
    CurvedConnector3 = 38,
    /// Curved connector 4
    CurvedConnector4 = 39,
    /// Curved connector 5
    CurvedConnector5 = 40,
    /// Callout 1
    Callout1 = 41,
    /// Callout 2
    Callout2 = 42,
    /// Callout 3
    Callout3 = 43,
    /// Accent callout 1
    AccentCallout1 = 44,
    /// Accent callout 2
    AccentCallout2 = 45,
    /// Accent callout 3
    AccentCallout3 = 46,
    /// Border callout 1
    BorderCallout1 = 47,
    /// Border callout 2
    BorderCallout2 = 48,
    /// Border callout 3
    BorderCallout3 = 49,
    /// Accent border callout 1
    AccentBorderCallout1 = 50,
    /// Accent border callout 2
    AccentBorderCallout2 = 51,
    /// Accent border callout 3
    AccentBorderCallout3 = 52,
    /// Custom shape
    Custom = 255,
}

impl From<u16> for EscherShapeType {
    fn from(value: u16) -> Self {
        match value {
            0 => EscherShapeType::NotPrimitive,
            1 => EscherShapeType::Rectangle,
            2 => EscherShapeType::RoundRectangle,
            3 => EscherShapeType::Oval,
            4 => EscherShapeType::Diamond,
            5 => EscherShapeType::Triangle,
            6 => EscherShapeType::RightTriangle,
            7 => EscherShapeType::Parallelogram,
            8 => EscherShapeType::Trapezoid,
            9 => EscherShapeType::Hexagon,
            10 => EscherShapeType::Octagon,
            11 => EscherShapeType::Plus,
            12 => EscherShapeType::Star,
            13 => EscherShapeType::Arrow,
            14 => EscherShapeType::ThickArrow,
            15 => EscherShapeType::HomePlate,
            16 => EscherShapeType::Cube,
            17 => EscherShapeType::Balloon,
            18 => EscherShapeType::Seal,
            19 => EscherShapeType::Arc,
            20 => EscherShapeType::Line,
            21 => EscherShapeType::Plaque,
            22 => EscherShapeType::Can,
            23 => EscherShapeType::Donut,
            24 => EscherShapeType::TextSimple,
            25 => EscherShapeType::TextOctagon,
            26 => EscherShapeType::TextHexagon,
            27 => EscherShapeType::TextCurve,
            28 => EscherShapeType::TextWave,
            29 => EscherShapeType::TextRing,
            30 => EscherShapeType::TextOnCurve,
            31 => EscherShapeType::TextOnRing,
            32 => EscherShapeType::StraightConnector1,
            33 => EscherShapeType::BentConnector2,
            34 => EscherShapeType::BentConnector3,
            35 => EscherShapeType::BentConnector4,
            36 => EscherShapeType::BentConnector5,
            37 => EscherShapeType::CurvedConnector2,
            38 => EscherShapeType::CurvedConnector3,
            39 => EscherShapeType::CurvedConnector4,
            40 => EscherShapeType::CurvedConnector5,
            41 => EscherShapeType::Callout1,
            42 => EscherShapeType::Callout2,
            43 => EscherShapeType::Callout3,
            44 => EscherShapeType::AccentCallout1,
            45 => EscherShapeType::AccentCallout2,
            46 => EscherShapeType::AccentCallout3,
            47 => EscherShapeType::BorderCallout1,
            48 => EscherShapeType::BorderCallout2,
            49 => EscherShapeType::BorderCallout3,
            50 => EscherShapeType::AccentBorderCallout1,
            51 => EscherShapeType::AccentBorderCallout2,
            52 => EscherShapeType::AccentBorderCallout3,
            255 => EscherShapeType::Custom,
            _ => EscherShapeType::NotPrimitive,
        }
    }
}

impl EscherShapeType {
    /// Get the u16 value of this shape type
    pub fn as_u16(self) -> u16 {
        unsafe { std::mem::transmute::<Self, u16>(self) }
    }
}
