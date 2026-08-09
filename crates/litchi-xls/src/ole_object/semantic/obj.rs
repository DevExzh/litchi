//! Semantic metadata for BIFF `Obj` records and embedded OLE objects.

use super::control::{FtCblsData, FtEdoData, FtGboData, FtLbsData, FtRboData, FtSbs};

/// The common object metadata subrecord (`FtCmo`, MS-XLS 2.5.143).
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct FtCmo {
    pub object_type: u16,
    pub object_id: u16,
    pub flags: u16,
    pub reserved: [u8; 12],
}

/// Picture/OLE flags (`FtPioGrbit`, MS-XLS 2.5.150).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct FtPioGrbit {
    /// The original bitfield, including currently undefined bits.
    pub raw: u16,
}

impl FtPioGrbit {
    #[must_use]
    pub fn is_dde(self) -> bool {
        self.raw & 2 != 0
    }

    #[must_use]
    pub fn display_as_icon(self) -> bool {
        self.raw & 8 != 0
    }

    #[must_use]
    pub fn is_control(self) -> bool {
        self.raw & 0x10 != 0
    }

    #[must_use]
    pub fn uses_control_stream(self) -> bool {
        self.raw & 0x20 != 0
    }

    #[must_use]
    pub fn camera_picture(self) -> bool {
        self.raw & 0x80 != 0
    }

    #[must_use]
    pub fn default_size(self) -> bool {
        self.raw & 0x100 != 0
    }

    #[must_use]
    pub fn auto_load(self) -> bool {
        self.raw & 0x200 != 0
    }
}

/// OLE picture formula metadata (`FtPictFmla`, MS-XLS 2.5.151).
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct FtPictFmla {
    pub formula: Vec<u8>,
    pub storage_position: Option<u32>,
    pub control_buffer_size: Option<u32>,
}

/// Type of object represented by an Obj record (`cmo.ot`, MS-XLS 2.5.143).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ObjectType {
    Group,
    Line,
    Rectangle,
    Oval,
    Arc,
    Chart,
    Text,
    Button,
    Picture,
    Polygon,
    CheckBox,
    RadioButton,
    EditBox,
    Label,
    DialogBox,
    SpinControl,
    ScrollBar,
    List,
    GroupBox,
    DropDown,
    Note,
    OfficeArt,
}

impl ObjectType {
    /// Whether this object type is represented by the form-control view.
    #[must_use]
    pub const fn is_form_control(self) -> bool {
        matches!(
            self,
            Self::CheckBox
                | Self::RadioButton
                | Self::EditBox
                | Self::Label
                | Self::DialogBox
                | Self::SpinControl
                | Self::ScrollBar
                | Self::List
                | Self::GroupBox
                | Self::DropDown
        )
    }

    /// Returns the MS-XLS `cmo.ot` code for this object type.
    #[must_use]
    pub const fn code(self) -> u16 {
        match self {
            Self::Group => 0x0000,
            Self::Line => 0x0001,
            Self::Rectangle => 0x0002,
            Self::Oval => 0x0003,
            Self::Arc => 0x0004,
            Self::Chart => 0x0005,
            Self::Text => 0x0006,
            Self::Button => 0x0007,
            Self::Picture => 0x0008,
            Self::Polygon => 0x0009,
            Self::CheckBox => 0x000B,
            Self::RadioButton => 0x000C,
            Self::EditBox => 0x000D,
            Self::Label => 0x000E,
            Self::DialogBox => 0x000F,
            Self::SpinControl => 0x0010,
            Self::ScrollBar => 0x0011,
            Self::List => 0x0012,
            Self::GroupBox => 0x0013,
            Self::DropDown => 0x0014,
            Self::Note => 0x0019,
            Self::OfficeArt => 0x001E,
        }
    }

    pub(crate) fn from_code(value: u16) -> Option<Self> {
        Some(match value {
            0x0000 => Self::Group,
            0x0001 => Self::Line,
            0x0002 => Self::Rectangle,
            0x0003 => Self::Oval,
            0x0004 => Self::Arc,
            0x0005 => Self::Chart,
            0x0006 => Self::Text,
            0x0007 => Self::Button,
            0x0008 => Self::Picture,
            0x0009 => Self::Polygon,
            0x000B => Self::CheckBox,
            0x000C => Self::RadioButton,
            0x000D => Self::EditBox,
            0x000E => Self::Label,
            0x000F => Self::DialogBox,
            0x0010 => Self::SpinControl,
            0x0011 => Self::ScrollBar,
            0x0012 => Self::List,
            0x0013 => Self::GroupBox,
            0x0014 => Self::DropDown,
            0x0019 => Self::Note,
            0x001E => Self::OfficeArt,
            _ => return None,
        })
    }
}

impl FtCmo {
    /// The object type decoded per MS-XLS 2.5.143, or `None` for an undefined
    /// value.
    #[must_use]
    pub fn object_kind(&self) -> Option<ObjectType> {
        ObjectType::from_code(self.object_type)
    }
}

/// A typed BIFF subrecord belonging to an Obj record.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum ObjSubrecord {
    Common(FtCmo),
    ClipboardFormat(Vec<u8>),
    PictureFlags(FtPioGrbit),
    PictureFormula(FtPictFmla),
    CheckBoxData(FtCblsData),
    RadioButtonData(FtRboData),
    EditBoxData(FtEdoData),
    GroupBoxData(FtGboData),
    ScrollBarData(FtSbs),
    ListBoxData(FtLbsData),
    /// A record kind not modeled by this crate; its bytes remain untouched.
    Unknown {
        kind: u16,
        data: Vec<u8>,
    },
    End,
}

/// A complete embedded-OLE Obj record and its adjacent `TxO` record.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct OleObjectRecord {
    pub subrecords: Vec<ObjSubrecord>,
    /// Complete adjacent `TxO` record, retained byte-for-byte.
    pub text_object: Option<Vec<u8>>,
}

impl OleObjectRecord {
    #[must_use]
    pub fn object_id(&self) -> u16 {
        self.subrecords
            .iter()
            .find_map(|value| match value {
                ObjSubrecord::Common(value) => Some(value.object_id),
                _ => None,
            })
            .unwrap_or(0)
    }

    #[must_use]
    pub fn storage_position(&self) -> Option<u32> {
        self.subrecords.iter().find_map(|value| match value {
            ObjSubrecord::PictureFormula(value) => value.storage_position,
            _ => None,
        })
    }

    #[must_use]
    pub fn storage_name(&self) -> Option<String> {
        let position = self.storage_position()?;
        let dde = self
            .subrecords
            .iter()
            .find_map(|value| match value {
                ObjSubrecord::PictureFlags(value) => Some(value.is_dde()),
                _ => None,
            })
            .unwrap_or(false);
        Some(format!(
            "{}{:08X}",
            if dde { "LNK" } else { "MBD" },
            position
        ))
    }
}
