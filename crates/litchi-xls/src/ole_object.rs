//! Strict BIFF8 Obj/FtPictFmla parsing and transactional OLE-object editing.

use super::{XlsError, XlsResult};
use litchi_cfb::OleFile;
use litchi_ole_common::object::{Editor as ObjectEditor, Target, Targets};
use std::collections::{HashMap, HashSet};
use std::io::Cursor;

pub use litchi_ole_common::object::Limits;

const OBJ: u16 = 0x005D;
const TXO: u16 = 0x01B6;
const CONTINUE: u16 = 0x003C;
const BOUNDSHEET: u16 = 0x0085;
const EOF: u16 = 0x000A;
const FT_CMO: u16 = 0x0015;
const FT_CF: u16 = 0x0007;
const FT_PIO: u16 = 0x0008;
const FT_PICT_FMLA: u16 = 0x0009;
const FT_SBS: u16 = 0x000C;
const FT_GBO_DATA: u16 = 0x000F;
const FT_EDO_DATA: u16 = 0x0010;
const FT_RBO_DATA: u16 = 0x0011;
const FT_CBLS_DATA: u16 = 0x0012;
const FT_LBS_DATA: u16 = 0x0013;
const FT_END: u16 = 0;

/// MS-XLS 2.5.141/2.5.145 `fNo3d`: the control is drawn without 3-D effects.
const NO_3D: u16 = 0x0001;
/// MS-XLS 2.5.154 FtSbs flag bits.
const SBS_DRAW: u16 = 0x0001;
const SBS_DRAW_SLIDER_ONLY: u16 = 0x0002;
const SBS_TRACK_ELEVATOR: u16 = 0x0004;
const SBS_NO_3D: u16 = 0x0008;
/// MS-XLS 2.5.147 FtLbsData flag bits.
const LBS_USE_CB: u16 = 0x0001;
const LBS_VALID_PLEX: u16 = 0x0002;
const LBS_VALID_IDS: u16 = 0x0004;
const LBS_NO_3D: u16 = 0x0008;
const LBS_SELECTION_TYPE_SHIFT: u16 = 4;
const LBS_SELECTION_TYPE_MASK: u16 = 0x0003;
const LBS_BEHAVIOR_CLASS_SHIFT: u16 = 8;
/// MS-XLS 2.5.171 LbsDropData flag bits.
const DROP_STYLE_MASK: u16 = 0x0003;
const DROP_FILTERED: u16 = 0x0004;
/// MS-XLS 2.5.294 XLUnicodeString option bits.
const XL_STRING_HIGH_BYTE: u8 = 0x01;
const XL_STRING_EXT: u8 = 0x04;
const XL_STRING_RICH: u8 = 0x08;
/// Size in bytes of one formatting run in a rich XLUnicodeString.
const FORMATTING_RUN_SIZE: usize = 4;
/// `cmo.ot` values that identify worksheet form controls (MS-XLS 2.5.143).
const OBJECT_TYPE_CHECK_BOX: u16 = 0x000B;
const OBJECT_TYPE_DROP_DOWN: u16 = 0x0014;

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XlsFtCmo {
    pub object_type: u16,
    pub object_id: u16,
    pub flags: u16,
    pub reserved: [u8; 12],
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct XlsFtPioGrbit {
    pub raw: u16,
}

impl XlsFtPioGrbit {
    pub fn is_dde(self) -> bool {
        self.raw & 2 != 0
    }
    pub fn display_as_icon(self) -> bool {
        self.raw & 8 != 0
    }
    pub fn is_control(self) -> bool {
        self.raw & 0x10 != 0
    }
    pub fn uses_control_stream(self) -> bool {
        self.raw & 0x20 != 0
    }
    pub fn camera_picture(self) -> bool {
        self.raw & 0x80 != 0
    }
    pub fn default_size(self) -> bool {
        self.raw & 0x100 != 0
    }
    pub fn auto_load(self) -> bool {
        self.raw & 0x200 != 0
    }
    fn validate(self) -> XlsResult<()> {
        if self.is_dde() && self.is_control() {
            return Err(invalid(
                OBJ,
                "FtPioGrbit DDE and control flags are mutually exclusive",
            ));
        }
        Ok(())
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XlsFtPictFmla {
    pub formula: Vec<u8>,
    pub storage_position: Option<u32>,
    pub control_buffer_size: Option<u32>,
}

/// Type of object represented by an Obj record (MS-XLS 2.5.143 `cmo.ot`).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum XlsObjectType {
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

impl XlsObjectType {
    fn from_code(value: u16) -> Option<Self> {
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

impl XlsFtCmo {
    /// The object type decoded per MS-XLS 2.5.143, or `None` for a value the
    /// specification does not define.
    pub fn object_kind(&self) -> Option<XlsObjectType> {
        XlsObjectType::from_code(self.object_type)
    }
}

/// State of a checkbox or radio button control (MS-XLS 2.5.141 `fChecked`).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum XlsCheckState {
    Unchecked,
    Checked,
    Mixed,
}

impl XlsCheckState {
    fn from_code(value: u16) -> Option<Self> {
        Some(match value {
            0x0000 => Self::Unchecked,
            0x0001 => Self::Checked,
            0x0002 => Self::Mixed,
            _ => return None,
        })
    }
    fn code(self) -> u16 {
        match self {
            Self::Unchecked => 0x0000,
            Self::Checked => 0x0001,
            Self::Mixed => 0x0002,
        }
    }
}

/// Input data validation expected by an edit box (MS-XLS 2.5.144 `ivtEdit`).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum XlsEditBoxValidation {
    AnyString,
    Integer,
    Number,
    Reference,
    Formula,
}

impl XlsEditBoxValidation {
    fn from_code(value: u16) -> Option<Self> {
        Some(match value {
            0x0000 => Self::AnyString,
            0x0001 => Self::Integer,
            0x0002 => Self::Number,
            0x0003 => Self::Reference,
            0x0004 => Self::Formula,
            _ => return None,
        })
    }
    fn code(self) -> u16 {
        match self {
            Self::AnyString => 0x0000,
            Self::Integer => 0x0001,
            Self::Number => 0x0002,
            Self::Reference => 0x0003,
            Self::Formula => 0x0004,
        }
    }
}

/// Selection behavior of a list control (MS-XLS 2.5.147 `wListSelType`).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum XlsListSelectionType {
    /// Only one item can be selected.
    Single,
    /// Multiple items can be selected by clicking each item.
    Multi,
    /// Multiple items can be selected by CTRL-clicking each item.
    CtrlMulti,
    /// Value 3; MS-XLS does not define a meaning for it.
    Reserved,
}

/// Behavior class of a list control (MS-XLS 2.5.147 `lct`).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum XlsListBehaviorClass {
    /// Regular sheet dropdown control.
    Regular,
    PivotPageField,
    AutoFilter,
    AutoComplete,
    DataValidation,
    PivotField,
    TotalRow,
    /// A value MS-XLS does not define.
    Unknown(u8),
}

/// Visual style of a dropdown control (MS-XLS 2.5.171 `wStyle`).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum XlsDropDownStyle {
    Combo,
    ComboEdit,
    Simple,
    /// Value 3; MS-XLS does not define a meaning for it.
    Reserved,
}

/// FtCblsData (MS-XLS 2.5.141): checkbox or radio button properties.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XlsFtCblsData {
    pub state: XlsCheckState,
    /// Unicode character of the accelerator key; 0 means none.
    pub accelerator: u16,
    pub reserved: u16,
    /// Raw `fNo3d`/`unused` bitfield; the unused bits are preserved verbatim.
    pub flags: u16,
}

impl XlsFtCblsData {
    /// Whether the control is drawn without three-dimensional effects.
    pub fn no_3d(&self) -> bool {
        self.flags & NO_3D != 0
    }
}

/// FtGboData (MS-XLS 2.5.145): group box properties.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XlsFtGboData {
    /// Unicode character of the accelerator key; 0 means none.
    pub accelerator: u16,
    pub reserved: u16,
    /// Raw `fNo3d`/`unused2` bitfield; the unused bits are preserved verbatim.
    pub flags: u16,
}

impl XlsFtGboData {
    /// Whether the control is drawn without three-dimensional effects.
    pub fn no_3d(&self) -> bool {
        self.flags & NO_3D != 0
    }
}

/// FtEdoData (MS-XLS 2.5.144): edit box properties.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XlsFtEdoData {
    pub validation: XlsEditBoxValidation,
    pub multi_line: bool,
    pub vertical_scroll_bar: bool,
    /// Identifier of the associated list control; 0 means none.
    pub list_control_id: u16,
}

/// FtRboData (MS-XLS 2.5.153): radio button grouping.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XlsFtRboData {
    /// Identifier of the next radio button in the group; 0 means none.
    pub next_radio_button_id: u16,
    /// Whether this is the first radio button of its group.
    pub first_in_group: bool,
}

/// FtSbs (MS-XLS 2.5.154): scroll bar or spin control properties.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XlsFtSbs {
    /// `unused1`, preserved verbatim.
    pub reserved: [u8; 4],
    pub value: i16,
    pub minimum: i16,
    pub maximum: i16,
    pub increment: i16,
    pub page_increment: i16,
    /// Whether the control scrolls horizontally rather than vertically.
    pub horizontal: bool,
    /// Width of the scroll bar in pixels.
    pub scroll_width: i16,
    /// Raw `fDraw`/`fDrawSliderOnly`/`fTrackElevator`/`fNo3d` bitfield; the
    /// unused bits are preserved verbatim.
    pub flags: u16,
}

impl XlsFtSbs {
    /// Whether the control is displayed (`fDraw`).
    pub fn draw(&self) -> bool {
        self.flags & SBS_DRAW != 0
    }
    /// Whether only the slider portion is displayed (`fDrawSliderOnly`).
    pub fn draw_slider_only(&self) -> bool {
        self.flags & SBS_DRAW_SLIDER_ONLY != 0
    }
    /// Whether the control tracks drags of the scroll thumb (`fTrackElevator`).
    pub fn track_elevator(&self) -> bool {
        self.flags & SBS_TRACK_ELEVATOR != 0
    }
    /// Whether the control is drawn without three-dimensional effects.
    pub fn no_3d(&self) -> bool {
        self.flags & SBS_NO_3D != 0
    }
}

/// A list item string that retains its original XLUnicodeString encoding, so a
/// read-write round-trip stays byte-identical.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XlsLbsItem {
    text: String,
    encoded: Vec<u8>,
}

impl XlsLbsItem {
    fn parse(encoded: Vec<u8>) -> Option<Self> {
        let text = decode_xl_unicode_string(&encoded)?;
        Some(Self { text, encoded })
    }

    /// Create an item from text, using compressed Unicode when every character
    /// fits in one byte. Returns `None` when the text exceeds the
    /// XLUnicodeString length limit.
    pub fn new(text: &str) -> Option<Self> {
        let mut encoded = Vec::new();
        if text.chars().all(|value| u32::from(value) <= 0xFF) {
            let count = u16::try_from(text.chars().count()).ok()?;
            encoded.extend_from_slice(&count.to_le_bytes());
            encoded.push(0);
            encoded.extend(text.chars().map(|value| value as u8));
        } else {
            let units = text.encode_utf16().collect::<Vec<_>>();
            let count = u16::try_from(units.len()).ok()?;
            encoded.extend_from_slice(&count.to_le_bytes());
            encoded.push(XL_STRING_HIGH_BYTE);
            encoded.extend(units.iter().flat_map(|unit| unit.to_le_bytes()));
        }
        Some(Self {
            text: text.to_string(),
            encoded,
        })
    }

    /// The item's text.
    pub fn text(&self) -> &str {
        &self.text
    }

    /// The original encoded XLUnicodeString, emitted verbatim on write.
    pub fn encoded(&self) -> &[u8] {
        &self.encoded
    }
}

/// LbsDropData (MS-XLS 2.5.171): dropdown-specific list box properties.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XlsLbsDropData {
    /// Raw `wStyle`/`fFiltered` bitfield; the unused bits are preserved verbatim.
    pub flags: u16,
    /// Number of lines displayed in the dropdown (`cLine`).
    pub line_count: u16,
    /// Smallest width in pixels allowed for the dropdown window (`dxMin`).
    pub min_width: u16,
    /// Current string value of the dropdown (`str`).
    text: XlsLbsItem,
    /// Trailing undefined byte, present iff `str` occupies an odd number of
    /// bytes (`unused3`); preserved verbatim.
    pub padding: Option<u8>,
}

impl XlsLbsDropData {
    /// The dropdown's visual style (`wStyle`).
    pub fn style(&self) -> XlsDropDownStyle {
        match self.flags & DROP_STYLE_MASK {
            0 => XlsDropDownStyle::Combo,
            1 => XlsDropDownStyle::ComboEdit,
            2 => XlsDropDownStyle::Simple,
            _ => XlsDropDownStyle::Reserved,
        }
    }
    /// Whether the displayed data has been filtered (`fFiltered`).
    pub fn filtered(&self) -> bool {
        self.flags & DROP_FILTERED != 0
    }
    /// The current string value of the dropdown.
    pub fn text(&self) -> &str {
        self.text.text()
    }
}

/// FtLbsData (MS-XLS 2.5.147): list box or dropdown properties.
///
/// When the record is continued into one or more Continue records, only the
/// portion present in the Obj record itself is typed; `items` and
/// `multi_selection` then hold fewer than `entry_count` elements.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct XlsFtLbsData {
    /// Raw ObjFmla payload (the `fmla` bytes following the `cbFmla` prefix)
    /// naming the range that fills the list.
    pub formula: Vec<u8>,
    /// Number of items in the list (`cLines`).
    pub entry_count: u16,
    /// One-based index of the first selected item (`iSel`); 0 means none.
    pub selected_index: u16,
    /// Raw flag/`lct` bitfield; the unused bits are preserved verbatim.
    pub flags: u16,
    /// Identifier of the associated edit box (`idEdit`); 0 means none.
    pub edit_box_id: u16,
    /// Dropdown properties; present iff the containing Obj is a dropdown.
    pub drop_down: Option<XlsLbsDropData>,
    /// List item strings (`rgLines`), each in its original encoding.
    items: Vec<XlsLbsItem>,
    /// Per-item multiple-selection state (`bsels`).
    pub multi_selection: Vec<bool>,
    /// Bytes following the typed portion, preserved verbatim on write.
    pub trailing: Vec<u8>,
}

impl XlsFtLbsData {
    /// Whether the `lct` behavior class is meaningful (`fUseCB`).
    pub fn has_behavior_class(&self) -> bool {
        self.flags & LBS_USE_CB != 0
    }
    /// Whether item strings are present (`fValidPlex`).
    pub fn has_item_strings(&self) -> bool {
        self.flags & LBS_VALID_PLEX != 0
    }
    /// Whether the edit box identifier is meaningful (`fValidIds`).
    pub fn has_edit_box(&self) -> bool {
        self.flags & LBS_VALID_IDS != 0
    }
    /// Whether the control is drawn without three-dimensional effects.
    pub fn no_3d(&self) -> bool {
        self.flags & LBS_NO_3D != 0
    }
    /// The selection behavior of the list control (`wListSelType`).
    pub fn selection_type(&self) -> XlsListSelectionType {
        match (self.flags >> LBS_SELECTION_TYPE_SHIFT) & LBS_SELECTION_TYPE_MASK {
            0 => XlsListSelectionType::Single,
            1 => XlsListSelectionType::Multi,
            2 => XlsListSelectionType::CtrlMulti,
            _ => XlsListSelectionType::Reserved,
        }
    }
    /// The behavior class of the list (`lct`).
    pub fn behavior_class(&self) -> XlsListBehaviorClass {
        match (self.flags >> LBS_BEHAVIOR_CLASS_SHIFT) as u8 {
            0x00 => XlsListBehaviorClass::Regular,
            0x01 => XlsListBehaviorClass::PivotPageField,
            0x03 => XlsListBehaviorClass::AutoFilter,
            0x05 => XlsListBehaviorClass::AutoComplete,
            0x06 => XlsListBehaviorClass::DataValidation,
            0x07 => XlsListBehaviorClass::PivotField,
            0x09 => XlsListBehaviorClass::TotalRow,
            value => XlsListBehaviorClass::Unknown(value),
        }
    }
    /// The list item strings.
    pub fn items(&self) -> &[XlsLbsItem] {
        &self.items
    }
    /// Replace the list item strings.
    pub fn set_items(&mut self, items: Vec<XlsLbsItem>) {
        self.items = items;
    }

    fn is_vacant(&self) -> bool {
        self.formula.is_empty()
            && self.entry_count == 0
            && self.selected_index == 0
            && self.flags == 0
            && self.edit_box_id == 0
            && self.drop_down.is_none()
            && self.items.is_empty()
            && self.multi_selection.is_empty()
            && self.trailing.is_empty()
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum XlsObjSubrecord {
    Common(XlsFtCmo),
    ClipboardFormat(Vec<u8>),
    PictureFlags(XlsFtPioGrbit),
    PictureFormula(XlsFtPictFmla),
    CheckBoxData(XlsFtCblsData),
    RadioButtonData(XlsFtRboData),
    EditBoxData(XlsFtEdoData),
    GroupBoxData(XlsFtGboData),
    ScrollBarData(XlsFtSbs),
    ListBoxData(XlsFtLbsData),
    Unknown { kind: u16, data: Vec<u8> },
    End,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XlsOleObjectRecord {
    pub subrecords: Vec<XlsObjSubrecord>,
    /// Complete adjacent TxO record, retained byte-for-byte.
    pub text_object: Option<Vec<u8>>,
}

impl XlsOleObjectRecord {
    pub fn parse(data: &[u8], text_object: Option<Vec<u8>>) -> XlsResult<Self> {
        let value = Self {
            subrecords: parse_subrecords(data)?,
            text_object,
        };
        value.validate()?;
        Ok(value)
    }

    pub fn object_id(&self) -> u16 {
        self.subrecords
            .iter()
            .find_map(|value| match value {
                XlsObjSubrecord::Common(value) => Some(value.object_id),
                _ => None,
            })
            .unwrap_or(0)
    }

    pub fn storage_position(&self) -> Option<u32> {
        self.subrecords.iter().find_map(|value| match value {
            XlsObjSubrecord::PictureFormula(value) => value.storage_position,
            _ => None,
        })
    }

    pub fn storage_name(&self) -> Option<String> {
        let position = self.storage_position()?;
        let dde = self
            .subrecords
            .iter()
            .find_map(|value| match value {
                XlsObjSubrecord::PictureFlags(value) => Some(value.is_dde()),
                _ => None,
            })
            .unwrap_or(false);
        Some(format!(
            "{}{:08X}",
            if dde { "LNK" } else { "MBD" },
            position
        ))
    }

    pub fn validate(&self) -> XlsResult<()> {
        if self.subrecords.len() > 1_024 {
            return Err(invalid(OBJ, "too many Obj subrecords"));
        }
        let common = self
            .subrecords
            .iter()
            .filter_map(|value| match value {
                XlsObjSubrecord::Common(value) => Some(value),
                _ => None,
            })
            .collect::<Vec<_>>();
        if common.len() != 1
            || common[0].object_type != 8
            || common[0].object_id == 0
            || !matches!(self.subrecords.first(), Some(XlsObjSubrecord::Common(_)))
        {
            return Err(invalid(
                OBJ,
                "OLE Obj requires a leading FtCmo type 8 with nonzero ID",
            ));
        }
        let pio = self
            .subrecords
            .iter()
            .filter_map(|value| match value {
                XlsObjSubrecord::PictureFlags(value) => Some(*value),
                _ => None,
            })
            .collect::<Vec<_>>();
        if pio.len() != 1 {
            return Err(invalid(OBJ, "OLE Obj requires one FtPioGrbit"));
        }
        pio[0].validate()?;
        if self
            .subrecords
            .iter()
            .filter(|value| matches!(value, XlsObjSubrecord::PictureFormula(_)))
            .count()
            > 1
        {
            return Err(invalid(OBJ, "duplicate FtPictFmla"));
        }
        if !matches!(self.subrecords.last(), Some(XlsObjSubrecord::End)) {
            return Err(invalid(OBJ, "OLE Obj must end with FtEnd"));
        }
        Ok(())
    }

    pub fn to_record_bytes(&self) -> XlsResult<Vec<u8>> {
        self.validate()?;
        record(OBJ, &serialize_subrecords(&self.subrecords)?)
    }
}

/// A worksheet form control (checkbox, radio button, edit box, group box,
/// spin/scroll bar, list box, or dropdown) backed by an Obj record.
///
/// Unlike [`XlsOleObjectRecord`], which enforces the strict OLE-object shape,
/// this view accepts any Obj whose `cmo.ot` names a form control and types its
/// control-specific subrecords per MS-XLS 2.5.141-2.5.154.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XlsFormControl {
    /// Complete Obj subrecord list; the first entry is always
    /// [`XlsObjSubrecord::Common`].
    pub subrecords: Vec<XlsObjSubrecord>,
    /// Complete adjacent TxO record, retained byte-for-byte.
    pub text_object: Option<Vec<u8>>,
}

impl XlsFormControl {
    /// Parse a form-control Obj record body. Returns `None` when the record is
    /// not a worksheet form control (picture, chart, note, ...) or its framing
    /// is broken; such records keep flowing through untouched.
    pub fn parse(data: &[u8], text_object: Option<Vec<u8>>) -> Option<Self> {
        let subrecords = parse_subrecords(data).ok()?;
        let common = match subrecords.first() {
            Some(XlsObjSubrecord::Common(value)) => value,
            _ => return None,
        };
        let control = (OBJECT_TYPE_CHECK_BOX..=OBJECT_TYPE_DROP_DOWN).contains(&common.object_type);
        control.then_some(Self {
            subrecords,
            text_object,
        })
    }

    /// The control's object identifier (`cmo.id`).
    pub fn object_id(&self) -> u16 {
        self.common().map_or(0, |value| value.object_id)
    }

    /// The kind of form control, decoded from `cmo.ot`.
    pub fn control_type(&self) -> Option<XlsObjectType> {
        self.common().and_then(XlsFtCmo::object_kind)
    }

    /// Checkbox or radio button state, when present.
    pub fn check_box_data(&self) -> Option<&XlsFtCblsData> {
        self.find(|value| match value {
            XlsObjSubrecord::CheckBoxData(value) => Some(value),
            _ => None,
        })
    }

    /// Radio button grouping, when present.
    pub fn radio_button_data(&self) -> Option<&XlsFtRboData> {
        self.find(|value| match value {
            XlsObjSubrecord::RadioButtonData(value) => Some(value),
            _ => None,
        })
    }

    /// Edit box properties, when present.
    pub fn edit_box_data(&self) -> Option<&XlsFtEdoData> {
        self.find(|value| match value {
            XlsObjSubrecord::EditBoxData(value) => Some(value),
            _ => None,
        })
    }

    /// Group box properties, when present.
    pub fn group_box_data(&self) -> Option<&XlsFtGboData> {
        self.find(|value| match value {
            XlsObjSubrecord::GroupBoxData(value) => Some(value),
            _ => None,
        })
    }

    /// Scroll bar or spin control properties, when present.
    pub fn scroll_bar_data(&self) -> Option<&XlsFtSbs> {
        self.find(|value| match value {
            XlsObjSubrecord::ScrollBarData(value) => Some(value),
            _ => None,
        })
    }

    /// List box or dropdown properties, when present.
    pub fn list_box_data(&self) -> Option<&XlsFtLbsData> {
        self.find(|value| match value {
            XlsObjSubrecord::ListBoxData(value) => Some(value),
            _ => None,
        })
    }

    /// Serialize the control back to a complete OBJ record.
    pub fn to_record_bytes(&self) -> XlsResult<Vec<u8>> {
        record(OBJ, &serialize_subrecords(&self.subrecords)?)
    }

    fn common(&self) -> Option<&XlsFtCmo> {
        self.find(|value| match value {
            XlsObjSubrecord::Common(value) => Some(value),
            _ => None,
        })
    }

    fn find<'a, T>(&'a self, pick: impl Fn(&'a XlsObjSubrecord) -> Option<&'a T>) -> Option<&'a T> {
        self.subrecords.iter().find_map(pick)
    }
}

#[derive(Clone)]
pub struct XlsOleObjectEditor {
    package: ObjectEditor,
    workbook_path: Vec<String>,
    workbook: Vec<u8>,
    sheets: Vec<Vec<XlsOleObjectRecord>>,
    form_controls: Vec<Vec<XlsFormControl>>,
}

impl XlsOleObjectEditor {
    pub fn new(bytes: Vec<u8>, limits: Limits) -> XlsResult<Self> {
        // Workbook metadata is XLS-owned. Read and parse it before handing
        // the original CFB bytes to the neutral object editor so the target
        // catalog can be derived solely from Obj/FtPictFmla records.
        let (workbook_path, workbook) = read_workbook(&bytes)?;
        let (sheets, form_controls) = parse_workbook(&workbook)?;
        let targets = targets_for_sheets(&sheets)?;
        let package = ObjectEditor::open(bytes, targets, limits)?;
        Ok(Self {
            package,
            workbook_path,
            workbook,
            sheets,
            form_controls,
        })
    }

    pub fn objects(&self, worksheet: usize) -> XlsResult<&[XlsOleObjectRecord]> {
        self.sheets
            .get(worksheet)
            .map(Vec::as_slice)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet index {worksheet}")))
    }

    /// Form controls (checkboxes, list boxes, scroll bars, ...) anchored in a
    /// worksheet, in Obj record order.
    pub fn form_controls(&self, worksheet: usize) -> XlsResult<&[XlsFormControl]> {
        self.form_controls
            .get(worksheet)
            .map(Vec::as_slice)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet index {worksheet}")))
    }

    pub fn add(
        &mut self,
        worksheet: usize,
        object: XlsOleObjectRecord,
        compound_file: Vec<u8>,
    ) -> XlsResult<()> {
        object.validate()?;
        let storage = object
            .storage_name()
            .ok_or_else(|| invalid(OBJ, "new Obj has no MBD/LNK reference"))?;
        if self
            .sheets
            .iter()
            .flatten()
            .any(|value| value.object_id() == object.object_id())
        {
            return Err(invalid(OBJ, "duplicate workbook object ID"));
        }
        let mut candidate = self.clone();
        candidate
            .sheets
            .get_mut(worksheet)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet index {worksheet}")))?
            .push(object);
        let target = target_for_storage(storage)?;
        candidate.package.add_storage(target, compound_file)?;
        candidate.commit()?;
        *self = candidate;
        Ok(())
    }

    pub fn remove(&mut self, worksheet: usize, object_id: u16) -> XlsResult<XlsOleObjectRecord> {
        let mut candidate = self.clone();
        let sheet = candidate
            .sheets
            .get_mut(worksheet)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet index {worksheet}")))?;
        let index = sheet
            .iter()
            .position(|value| value.object_id() == object_id)
            .ok_or_else(|| invalid(OBJ, "OLE object ID not found"))?;
        let removed = sheet.remove(index);
        if let Some(storage) = removed.storage_name()
            && !candidate
                .sheets
                .iter()
                .flatten()
                .any(|value| value.storage_name().as_deref() == Some(&storage))
        {
            let target = target_for_storage(storage)?;
            candidate.package.remove_storage(target.key())?;
        }
        candidate.commit()?;
        *self = candidate;
        Ok(removed)
    }

    pub fn reorder(&mut self, worksheet: usize, ids: &[u16]) -> XlsResult<()> {
        let mut candidate = self.clone();
        let sheet = candidate
            .sheets
            .get_mut(worksheet)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet index {worksheet}")))?;
        if ids.len() != sheet.len() {
            return Err(invalid(
                OBJ,
                "reorder must contain every worksheet OLE object",
            ));
        }
        let mut remaining = sheet.clone();
        let mut reordered = Vec::with_capacity(ids.len());
        for id in ids {
            let index = remaining
                .iter()
                .position(|value| value.object_id() == *id)
                .ok_or_else(|| invalid(OBJ, "unknown or repeated OLE object ID"))?;
            reordered.push(remaining.remove(index));
        }
        *sheet = reordered;
        candidate.commit()?;
        *self = candidate;
        Ok(())
    }

    pub fn replace_storage(&mut self, storage_name: &str, compound_file: Vec<u8>) -> XlsResult<()> {
        let storage = self
            .sheets
            .iter()
            .flatten()
            .find_map(|value| {
                value
                    .storage_name()
                    .filter(|value| value.as_str() == storage_name)
            })
            .ok_or_else(|| invalid(OBJ, "storage has no Obj reference"))?;
        let target = target_for_storage(storage)?;
        self.package
            .replace(target.key(), compound_file)
            .map_err(Into::into)
    }

    pub fn finish(self) -> XlsResult<Vec<u8>> {
        self.package.finish().map_err(Into::into)
    }

    fn commit(&mut self) -> XlsResult<()> {
        validate_objects(&self.sheets)?;
        let workbook = rewrite_workbook(&self.workbook, &self.sheets)?;
        self.package
            .put_stream(&self.workbook_path, workbook.clone())?;
        self.workbook = workbook;
        Ok(())
    }
}

fn read_workbook(bytes: &[u8]) -> XlsResult<(Vec<String>, Vec<u8>)> {
    let mut ole = OleFile::open(Cursor::new(bytes))?;
    for name in ["Workbook", "Book"] {
        match ole.open_stream(&[name]) {
            Ok(workbook) => return Ok((vec![name.to_owned()], workbook)),
            Err(litchi_cfb::OleError::StreamNotFound) => continue,
            Err(error) => return Err(error.into()),
        }
    }
    Err(XlsError::InvalidData("Workbook stream not found".into()))
}

fn target_for_storage(storage: String) -> XlsResult<Target> {
    Ok(Target::new(storage.clone(), [storage])?)
}

fn targets_for_sheets(sheets: &[Vec<XlsOleObjectRecord>]) -> XlsResult<Targets> {
    let mut seen = HashSet::new();
    let mut targets = Vec::new();
    for object in sheets.iter().flatten() {
        let Some(storage) = object.storage_name() else {
            continue;
        };
        if seen.insert(storage.clone()) {
            targets.push(target_for_storage(storage)?);
        }
    }
    Ok(Targets::new(targets)?)
}

fn parse_formula(body: &[u8]) -> XlsResult<XlsFtPictFmla> {
    if body.len() < 2 {
        return Err(invalid(OBJ, "FtPictFmla is truncated"));
    }
    let len = usize::from(u16::from_le_bytes([body[0], body[1]]));
    let end = 2usize
        .checked_add(len)
        .ok_or_else(|| invalid(OBJ, "formula overflow"))?;
    let formula = body
        .get(2..end)
        .ok_or_else(|| invalid(OBJ, "formula is truncated"))?
        .to_vec();
    let tail = &body[end..];
    let (storage_position, control_buffer_size) = match tail.len() {
        0 => (None, None),
        8 => (
            Some(u32_at(tail, 0).ok_or_else(|| invalid(OBJ, "storage position is truncated"))?),
            Some(u32_at(tail, 4).ok_or_else(|| invalid(OBJ, "control buffer size is truncated"))?),
        ),
        _ => return Err(invalid(OBJ, "unsupported FtPictFmla trailing layout")),
    };
    Ok(XlsFtPictFmla {
        formula,
        storage_position,
        control_buffer_size,
    })
}

fn parse_subrecords(data: &[u8]) -> XlsResult<Vec<XlsObjSubrecord>> {
    let mut offset = 0usize;
    let mut control_type = None;
    let mut subrecords = Vec::new();
    while offset < data.len() {
        let header = data
            .get(offset..offset + 4)
            .ok_or_else(|| invalid(OBJ, "truncated Obj subrecord header"))?;
        let kind = u16::from_le_bytes([header[0], header[1]]);
        let len = usize::from(u16::from_le_bytes([header[2], header[3]]));
        offset += 4;
        let end = offset
            .checked_add(len)
            .ok_or_else(|| invalid(OBJ, "Obj subrecord overflow"))?;
        let body = data
            .get(offset..end)
            .ok_or_else(|| invalid(OBJ, "truncated Obj subrecord"))?;
        let value = match (kind, len) {
            (FT_CMO, 18) => XlsObjSubrecord::Common(XlsFtCmo {
                object_type: u16::from_le_bytes([body[0], body[1]]),
                object_id: u16::from_le_bytes([body[2], body[3]]),
                flags: u16::from_le_bytes([body[4], body[5]]),
                reserved: array_at(body, 6)
                    .ok_or_else(|| invalid(OBJ, "FtCmo reserved bytes are truncated"))?,
            }),
            (FT_CMO, _) => return Err(invalid(OBJ, "FtCmo must contain 18 bytes")),
            (FT_CF, _) => XlsObjSubrecord::ClipboardFormat(body.to_vec()),
            (FT_PIO, 2) => XlsObjSubrecord::PictureFlags(XlsFtPioGrbit {
                raw: u16::from_le_bytes([body[0], body[1]]),
            }),
            (FT_PIO, _) => return Err(invalid(OBJ, "FtPioGrbit must contain 2 bytes")),
            (FT_PICT_FMLA, _) => XlsObjSubrecord::PictureFormula(parse_formula(body)?),
            // Form-control data subrecords fall back to raw preservation when
            // their contents do not match the MS-XLS layout.
            (FT_CBLS_DATA, _) => parse_cbls_data(body)
                .map_or_else(|| unknown(kind, body), XlsObjSubrecord::CheckBoxData),
            (FT_RBO_DATA, _) => parse_rbo_data(body)
                .map_or_else(|| unknown(kind, body), XlsObjSubrecord::RadioButtonData),
            (FT_EDO_DATA, _) => parse_edo_data(body)
                .map_or_else(|| unknown(kind, body), XlsObjSubrecord::EditBoxData),
            (FT_GBO_DATA, _) => parse_gbo_data(body)
                .map_or_else(|| unknown(kind, body), XlsObjSubrecord::GroupBoxData),
            (FT_SBS, _) => {
                parse_sbs(body).map_or_else(|| unknown(kind, body), XlsObjSubrecord::ScrollBarData)
            },
            (FT_LBS_DATA, _) => parse_lbs_data(body, control_type)
                .map_or_else(|| unknown(kind, body), XlsObjSubrecord::ListBoxData),
            (FT_END, 0) => XlsObjSubrecord::End,
            (FT_END, _) => return Err(invalid(OBJ, "FtEnd must be empty")),
            _ => unknown(kind, body),
        };
        if let XlsObjSubrecord::Common(common) = &value {
            control_type = Some(common.object_type);
        }
        subrecords.push(value);
        offset = end;
    }
    Ok(subrecords)
}

fn unknown(kind: u16, body: &[u8]) -> XlsObjSubrecord {
    XlsObjSubrecord::Unknown {
        kind,
        data: body.to_vec(),
    }
}

fn u16_at(data: &[u8], offset: usize) -> Option<u16> {
    Some(u16::from_le_bytes(array_at(data, offset)?))
}

fn u32_at(data: &[u8], offset: usize) -> Option<u32> {
    Some(u32::from_le_bytes(array_at(data, offset)?))
}

fn array_at<const N: usize>(data: &[u8], offset: usize) -> Option<[u8; N]> {
    let end = offset.checked_add(N)?;
    data.get(offset..end)?.try_into().ok()
}

fn bool_at(data: &[u8], offset: usize) -> Option<bool> {
    match u16_at(data, offset)? {
        0 => Some(false),
        1 => Some(true),
        _ => None,
    }
}

fn parse_cbls_data(body: &[u8]) -> Option<XlsFtCblsData> {
    if body.len() != 8 {
        return None;
    }
    Some(XlsFtCblsData {
        state: XlsCheckState::from_code(u16_at(body, 0)?)?,
        accelerator: u16_at(body, 2)?,
        reserved: u16_at(body, 4)?,
        flags: u16_at(body, 6)?,
    })
}

fn parse_rbo_data(body: &[u8]) -> Option<XlsFtRboData> {
    if body.len() != 4 {
        return None;
    }
    Some(XlsFtRboData {
        next_radio_button_id: u16_at(body, 0)?,
        first_in_group: bool_at(body, 2)?,
    })
}

fn parse_edo_data(body: &[u8]) -> Option<XlsFtEdoData> {
    if body.len() != 8 {
        return None;
    }
    Some(XlsFtEdoData {
        validation: XlsEditBoxValidation::from_code(u16_at(body, 0)?)?,
        multi_line: bool_at(body, 2)?,
        vertical_scroll_bar: bool_at(body, 4)?,
        list_control_id: u16_at(body, 6)?,
    })
}

fn parse_gbo_data(body: &[u8]) -> Option<XlsFtGboData> {
    if body.len() != 6 {
        return None;
    }
    Some(XlsFtGboData {
        accelerator: u16_at(body, 0)?,
        reserved: u16_at(body, 2)?,
        flags: u16_at(body, 4)?,
    })
}

fn parse_sbs(body: &[u8]) -> Option<XlsFtSbs> {
    if body.len() != 20 {
        return None;
    }
    Some(XlsFtSbs {
        reserved: array_at(body, 0)?,
        value: i16::from_le_bytes(array_at(body, 4)?),
        minimum: i16::from_le_bytes(array_at(body, 6)?),
        maximum: i16::from_le_bytes(array_at(body, 8)?),
        increment: i16::from_le_bytes(array_at(body, 10)?),
        page_increment: i16::from_le_bytes(array_at(body, 12)?),
        horizontal: bool_at(body, 14)?,
        scroll_width: i16::from_le_bytes(array_at(body, 16)?),
        flags: u16_at(body, 18)?,
    })
}

fn parse_lbs_data(body: &[u8], control_type: Option<u16>) -> Option<XlsFtLbsData> {
    if body.is_empty() {
        return Some(XlsFtLbsData::default());
    }
    let formula_len = usize::from(u16_at(body, 0)?);
    let formula_end = 2usize.checked_add(formula_len)?;
    let formula = body.get(2..formula_end)?.to_vec();
    let header_end = formula_end.checked_add(8)?;
    if body.len() < header_end {
        return None;
    }
    let mut data = XlsFtLbsData {
        formula,
        entry_count: u16_at(body, formula_end)?,
        selected_index: u16_at(body, formula_end + 2)?,
        flags: u16_at(body, formula_end + 4)?,
        edit_box_id: u16_at(body, formula_end + 6)?,
        ..XlsFtLbsData::default()
    };
    let mut offset = header_end;
    if control_type == Some(OBJECT_TYPE_DROP_DOWN) {
        let drop_header_end = offset.checked_add(6)?;
        if body.len() < drop_header_end {
            return None;
        }
        let flags = u16_at(body, offset)?;
        let line_count = u16_at(body, offset + 2)?;
        let min_width = u16_at(body, offset + 4)?;
        offset = drop_header_end;
        let text_len = xl_unicode_string_size(body.get(offset..)?)?;
        let text = XlsLbsItem::parse(body.get(offset..offset + text_len)?.to_vec())?;
        offset += text_len;
        let padding = if text_len % 2 == 1 {
            let value = *body.get(offset)?;
            offset += 1;
            Some(value)
        } else {
            None
        };
        data.drop_down = Some(XlsLbsDropData {
            flags,
            line_count,
            min_width,
            text,
            padding,
        });
    }
    // rgLines: parse up to `entry_count` item strings. A record continued into
    // Continue records holds fewer strings here; a defective string stops the
    // walk and its bytes are preserved verbatim as trailing data.
    let mut items = Vec::new();
    while items.len() < usize::from(data.entry_count) && offset < body.len() {
        match xl_unicode_string_size(&body[offset..]) {
            Some(size) if offset + size <= body.len() => {
                items.push(XlsLbsItem::parse(body[offset..offset + size].to_vec())?);
                offset += size;
            },
            _ => break,
        }
    }
    if offset < body.len() {
        // bsels: one selection byte per entry for multiple-selection lists.
        let multiple = (data.flags >> LBS_SELECTION_TYPE_SHIFT) & LBS_SELECTION_TYPE_MASK != 0;
        if multiple {
            let count = usize::from(data.entry_count).min(body.len() - offset);
            data.multi_selection = body[offset..offset + count]
                .iter()
                .map(|value| *value != 0)
                .collect();
            offset += count;
        }
        data.trailing = body[offset..].to_vec();
    }
    data.set_items(items);
    Some(data)
}

/// Total byte size of the XLUnicodeString (MS-XLS 2.5.294) starting at
/// `data`, including formatting runs and extension data, or `None` when the
/// framing is truncated or inconsistent.
fn xl_unicode_string_size(data: &[u8]) -> Option<usize> {
    if data.len() < 3 {
        return None;
    }
    let character_count = usize::from(u16_at(data, 0)?);
    let options = *data.get(2)?;
    let mut offset = 3usize;
    let formatting_runs = if options & XL_STRING_RICH != 0 {
        let count = usize::from(u16_at(data, offset)?);
        offset += 2;
        count
    } else {
        0
    };
    let extension_size = if options & XL_STRING_EXT != 0 {
        let size = u32::from_le_bytes(data.get(offset..offset + 4)?.try_into().ok()?) as usize;
        offset += 4;
        size
    } else {
        0
    };
    let character_bytes = character_count.checked_mul(if options & XL_STRING_HIGH_BYTE != 0 {
        2
    } else {
        1
    })?;
    let total = offset
        .checked_add(character_bytes)?
        .checked_add(formatting_runs.checked_mul(FORMATTING_RUN_SIZE)?)?
        .checked_add(extension_size)?;
    if total > data.len() {
        return None;
    }
    Some(total)
}

/// Decode the text of an exact-size XLUnicodeString, ignoring formatting runs
/// and extension data.
fn decode_xl_unicode_string(encoded: &[u8]) -> Option<String> {
    if xl_unicode_string_size(encoded)? != encoded.len() {
        return None;
    }
    let character_count = usize::from(u16_at(encoded, 0)?);
    let options = *encoded.get(2)?;
    let mut offset = 3usize;
    if options & XL_STRING_RICH != 0 {
        offset += 2;
    }
    if options & XL_STRING_EXT != 0 {
        offset += 4;
    }
    if options & XL_STRING_HIGH_BYTE != 0 {
        let bytes = encoded.get(offset..offset + character_count * 2)?;
        let units = bytes
            .chunks_exact(2)
            .map(|pair| u16::from_le_bytes([pair[0], pair[1]]));
        String::from_utf16(&units.collect::<Vec<_>>()).ok()
    } else {
        let bytes = encoded.get(offset..offset + character_count)?;
        Some(bytes.iter().map(|value| char::from(*value)).collect())
    }
}

fn serialize_subrecords(subrecords: &[XlsObjSubrecord]) -> XlsResult<Vec<u8>> {
    let mut output = Vec::new();
    for value in subrecords {
        let (kind, body) = serialize_subrecord(value)?;
        let len =
            u16::try_from(body.len()).map_err(|_| invalid(kind, "Obj subrecord exceeds u16"))?;
        output.extend_from_slice(&kind.to_le_bytes());
        output.extend_from_slice(&len.to_le_bytes());
        output.extend_from_slice(&body);
    }
    Ok(output)
}

fn serialize_subrecord(value: &XlsObjSubrecord) -> XlsResult<(u16, Vec<u8>)> {
    Ok(match value {
        XlsObjSubrecord::Common(value) => {
            let mut body = Vec::with_capacity(18);
            body.extend_from_slice(&value.object_type.to_le_bytes());
            body.extend_from_slice(&value.object_id.to_le_bytes());
            body.extend_from_slice(&value.flags.to_le_bytes());
            body.extend_from_slice(&value.reserved);
            (FT_CMO, body)
        },
        XlsObjSubrecord::ClipboardFormat(data) => (FT_CF, data.clone()),
        XlsObjSubrecord::PictureFlags(value) => (FT_PIO, value.raw.to_le_bytes().to_vec()),
        XlsObjSubrecord::PictureFormula(value) => {
            let len = u16::try_from(value.formula.len())
                .map_err(|_| invalid(OBJ, "formula exceeds u16"))?;
            let mut body = len.to_le_bytes().to_vec();
            body.extend_from_slice(&value.formula);
            match (value.storage_position, value.control_buffer_size) {
                (Some(position), Some(size)) => {
                    body.extend_from_slice(&position.to_le_bytes());
                    body.extend_from_slice(&size.to_le_bytes());
                },
                (None, None) => {},
                _ => {
                    return Err(invalid(
                        OBJ,
                        "FtPictFmla optional fields must occur together",
                    ));
                },
            }
            (FT_PICT_FMLA, body)
        },
        XlsObjSubrecord::CheckBoxData(value) => {
            let mut body = Vec::with_capacity(8);
            body.extend_from_slice(&value.state.code().to_le_bytes());
            body.extend_from_slice(&value.accelerator.to_le_bytes());
            body.extend_from_slice(&value.reserved.to_le_bytes());
            body.extend_from_slice(&value.flags.to_le_bytes());
            (FT_CBLS_DATA, body)
        },
        XlsObjSubrecord::RadioButtonData(value) => {
            let mut body = Vec::with_capacity(4);
            body.extend_from_slice(&value.next_radio_button_id.to_le_bytes());
            body.extend_from_slice(&u16::from(value.first_in_group).to_le_bytes());
            (FT_RBO_DATA, body)
        },
        XlsObjSubrecord::EditBoxData(value) => {
            let mut body = Vec::with_capacity(8);
            body.extend_from_slice(&value.validation.code().to_le_bytes());
            body.extend_from_slice(&u16::from(value.multi_line).to_le_bytes());
            body.extend_from_slice(&u16::from(value.vertical_scroll_bar).to_le_bytes());
            body.extend_from_slice(&value.list_control_id.to_le_bytes());
            (FT_EDO_DATA, body)
        },
        XlsObjSubrecord::GroupBoxData(value) => {
            let mut body = Vec::with_capacity(6);
            body.extend_from_slice(&value.accelerator.to_le_bytes());
            body.extend_from_slice(&value.reserved.to_le_bytes());
            body.extend_from_slice(&value.flags.to_le_bytes());
            (FT_GBO_DATA, body)
        },
        XlsObjSubrecord::ScrollBarData(value) => {
            let mut body = Vec::with_capacity(20);
            body.extend_from_slice(&value.reserved);
            body.extend_from_slice(&value.value.to_le_bytes());
            body.extend_from_slice(&value.minimum.to_le_bytes());
            body.extend_from_slice(&value.maximum.to_le_bytes());
            body.extend_from_slice(&value.increment.to_le_bytes());
            body.extend_from_slice(&value.page_increment.to_le_bytes());
            body.extend_from_slice(&u16::from(value.horizontal).to_le_bytes());
            body.extend_from_slice(&value.scroll_width.to_le_bytes());
            body.extend_from_slice(&value.flags.to_le_bytes());
            (FT_SBS, body)
        },
        XlsObjSubrecord::ListBoxData(value) => {
            if value.is_vacant() {
                (FT_LBS_DATA, Vec::new())
            } else {
                let len = u16::try_from(value.formula.len())
                    .map_err(|_| invalid(OBJ, "ObjFmla exceeds u16"))?;
                let mut body = len.to_le_bytes().to_vec();
                body.extend_from_slice(&value.formula);
                body.extend_from_slice(&value.entry_count.to_le_bytes());
                body.extend_from_slice(&value.selected_index.to_le_bytes());
                body.extend_from_slice(&value.flags.to_le_bytes());
                body.extend_from_slice(&value.edit_box_id.to_le_bytes());
                if let Some(drop_down) = &value.drop_down {
                    body.extend_from_slice(&drop_down.flags.to_le_bytes());
                    body.extend_from_slice(&drop_down.line_count.to_le_bytes());
                    body.extend_from_slice(&drop_down.min_width.to_le_bytes());
                    body.extend_from_slice(&drop_down.text.encoded);
                    if drop_down.text.encoded.len() % 2 == 1 {
                        body.push(drop_down.padding.unwrap_or(0));
                    }
                }
                for item in value.items() {
                    body.extend_from_slice(item.encoded());
                }
                body.extend(
                    value
                        .multi_selection
                        .iter()
                        .map(|selected| u8::from(*selected)),
                );
                body.extend_from_slice(&value.trailing);
                (FT_LBS_DATA, body)
            }
        },
        XlsObjSubrecord::Unknown { kind, data } => (*kind, data.clone()),
        XlsObjSubrecord::End => (FT_END, Vec::new()),
    })
}

#[allow(clippy::type_complexity)]
fn parse_workbook(
    input: &[u8],
) -> XlsResult<(Vec<Vec<XlsOleObjectRecord>>, Vec<Vec<XlsFormControl>>)> {
    let (_, starts) = bindings(input)?;
    let mut sheets = Vec::new();
    let mut form_controls = Vec::new();
    for (index, (start, worksheet)) in starts.iter().enumerate() {
        if !worksheet {
            continue;
        }
        let end = starts.get(index + 1).map_or(input.len(), |value| value.0);
        let (objects, controls) = parse_sheet(&input[*start..end])?;
        sheets.push(objects);
        form_controls.push(controls);
    }
    validate_objects(&sheets)?;
    Ok((sheets, form_controls))
}

fn parse_sheet(input: &[u8]) -> XlsResult<(Vec<XlsOleObjectRecord>, Vec<XlsFormControl>)> {
    let records = ranges(input)?;
    let mut objects = Vec::new();
    let mut controls = Vec::new();
    for (index, value) in records.iter().enumerate() {
        if value.2 != OBJ {
            continue;
        }
        let txo = if records.get(index + 1).is_some_and(|next| next.2 == TXO) {
            if records
                .get(index + 2)
                .is_some_and(|next| next.2 == CONTINUE)
            {
                return Err(invalid(
                    TXO,
                    "Continue-based TxO beside OLE Obj is unsupported",
                ));
            }
            let next = records[index + 1];
            Some(input[next.0..next.1].to_vec())
        } else {
            None
        };
        let body = &input[value.3..value.4];
        if let Ok(object) = XlsOleObjectRecord::parse(body, txo.clone()) {
            objects.push(object);
        } else if let Some(control) = XlsFormControl::parse(body, txo) {
            controls.push(control);
        }
    }
    Ok((objects, controls))
}

fn validate_objects(sheets: &[Vec<XlsOleObjectRecord>]) -> XlsResult<()> {
    let mut ids = HashSet::new();
    for (index, object) in sheets.iter().flatten().enumerate() {
        if index >= 4_096 {
            return Err(invalid(OBJ, "workbook object count exceeds limit"));
        }
        object.validate()?;
        if !ids.insert(object.object_id()) {
            return Err(invalid(OBJ, "duplicate workbook object ID"));
        }
    }
    Ok(())
}

fn rewrite_workbook(input: &[u8], sheets: &[Vec<XlsOleObjectRecord>]) -> XlsResult<Vec<u8>> {
    let (refs, starts) = bindings(input)?;
    let first = starts.first().map_or(input.len(), |value| value.0);
    let mut output = input[..first].to_vec();
    let mut new_offsets = HashMap::new();
    let mut worksheet = 0usize;
    for (index, (start, is_worksheet)) in starts.iter().enumerate() {
        let end = starts.get(index + 1).map_or(input.len(), |value| value.0);
        new_offsets.insert(*start, output.len());
        if *is_worksheet {
            output.extend_from_slice(&rewrite_sheet(
                &input[*start..end],
                sheets
                    .get(worksheet)
                    .ok_or_else(|| invalid(BOUNDSHEET, "worksheet list missing"))?,
            )?);
            worksheet += 1;
        } else {
            output.extend_from_slice(&input[*start..end]);
        }
    }
    if worksheet != sheets.len() {
        return Err(invalid(BOUNDSHEET, "worksheet count mismatch"));
    }
    for (payload, old) in refs {
        let new = *new_offsets
            .get(&old)
            .ok_or_else(|| invalid(BOUNDSHEET, "sheet target missing"))?;
        output[payload..payload + 4].copy_from_slice(
            &u32::try_from(new)
                .map_err(|_| invalid(BOUNDSHEET, "sheet offset exceeds u32"))?
                .to_le_bytes(),
        );
    }
    Ok(output)
}

fn rewrite_sheet(input: &[u8], objects: &[XlsOleObjectRecord]) -> XlsResult<Vec<u8>> {
    let records = ranges(input)?;
    let mut output = Vec::new();
    let mut next = 0usize;
    let mut skip_txo = false;
    for (index, value) in records.iter().enumerate() {
        if skip_txo && value.2 == TXO {
            skip_txo = false;
            continue;
        }
        if value.2 == OBJ && XlsOleObjectRecord::parse(&input[value.3..value.4], None).is_ok() {
            if let Some(object) = objects.get(next) {
                output.extend_from_slice(&object.to_record_bytes()?);
                if let Some(txo) = &object.text_object {
                    output.extend_from_slice(txo);
                }
                next += 1;
            }
            skip_txo = records
                .get(index + 1)
                .is_some_and(|following| following.2 == TXO);
            continue;
        }
        if value.2 == EOF {
            for object in &objects[next..] {
                output.extend_from_slice(&object.to_record_bytes()?);
                if let Some(txo) = &object.text_object {
                    output.extend_from_slice(txo);
                }
            }
            next = objects.len();
        }
        output.extend_from_slice(&input[value.0..value.1]);
    }
    if next != objects.len() {
        return Err(invalid(EOF, "worksheet has no EOF"));
    }
    Ok(output)
}

#[allow(clippy::type_complexity)]
fn bindings(input: &[u8]) -> XlsResult<(Vec<(usize, usize)>, Vec<(usize, bool)>)> {
    let mut refs = Vec::new();
    for (start, _, kind, body_start, body_end) in ranges(input)? {
        if kind != BOUNDSHEET {
            continue;
        }
        let body = &input[body_start..body_end];
        if body.len() < 6 {
            return Err(invalid(BOUNDSHEET, "BoundSheet is truncated"));
        }
        refs.push((
            start
                .checked_add(4)
                .ok_or_else(|| invalid(BOUNDSHEET, "record offset overflow"))?,
            u32_at(body, 0).ok_or_else(|| invalid(BOUNDSHEET, "sheet offset is truncated"))?
                as usize,
            body[5] == 0,
        ));
    }
    let mut starts = refs
        .iter()
        .map(|(_, offset, sheet)| (*offset, *sheet))
        .collect::<Vec<_>>();
    starts.sort_by_key(|value| value.0);
    if starts.windows(2).any(|value| value[0].0 >= value[1].0)
        || starts.iter().any(|value| value.0 >= input.len())
    {
        return Err(invalid(BOUNDSHEET, "invalid or duplicate sheet offsets"));
    }
    Ok((
        refs.into_iter()
            .map(|(payload, offset, _)| (payload, offset))
            .collect(),
        starts,
    ))
}

#[allow(clippy::type_complexity)]
fn ranges(input: &[u8]) -> XlsResult<Vec<(usize, usize, u16, usize, usize)>> {
    let mut output = Vec::new();
    let mut offset = 0usize;
    while offset < input.len() {
        let header = input
            .get(offset..offset + 4)
            .ok_or(XlsError::InvalidLength {
                expected: offset + 4,
                found: input.len(),
            })?;
        let kind = u16::from_le_bytes([header[0], header[1]]);
        let len = usize::from(u16::from_le_bytes([header[2], header[3]]));
        let end = offset
            .checked_add(4 + len)
            .ok_or_else(|| invalid(kind, "record size overflow"))?;
        if end > input.len() {
            return Err(XlsError::InvalidLength {
                expected: end,
                found: input.len(),
            });
        }
        output.push((offset, end, kind, offset + 4, end));
        offset = end;
    }
    Ok(output)
}

fn record(kind: u16, body: &[u8]) -> XlsResult<Vec<u8>> {
    if body.len() > 8_224 {
        return Err(invalid(kind, "record exceeds BIFF8 limit"));
    }
    let mut output = Vec::with_capacity(body.len() + 4);
    output.extend_from_slice(&kind.to_le_bytes());
    output.extend_from_slice(&(body.len() as u16).to_le_bytes());
    output.extend_from_slice(body);
    Ok(output)
}

fn invalid(record_type: u16, message: impl Into<String>) -> XlsError {
    XlsError::InvalidRecord {
        record_type,
        message: message.into(),
    }
}
