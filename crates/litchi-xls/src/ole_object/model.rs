//! Typed Obj and worksheet-control models defined by MS-XLS.

use super::codec::{
    decode_xl_unicode_string, parse_subrecords, record, serialize_subrecords, u16_at,
};
use super::*;
use crate::error::Result;

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct FtCmo {
    pub object_type: u16,
    pub object_id: u16,
    pub flags: u16,
    pub reserved: [u8; 12],
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct FtPioGrbit {
    pub raw: u16,
}

impl FtPioGrbit {
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
    fn validate(self) -> Result<()> {
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
pub struct FtPictFmla {
    pub formula: Vec<u8>,
    pub storage_position: Option<u32>,
    pub control_buffer_size: Option<u32>,
}

/// Type of object represented by an Obj record (MS-XLS 2.5.143 `cmo.ot`).
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

    pub(super) fn from_code(value: u16) -> Option<Self> {
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
    /// The object type decoded per MS-XLS 2.5.143, or `None` for a value the
    /// specification does not define.
    pub fn object_kind(&self) -> Option<ObjectType> {
        ObjectType::from_code(self.object_type)
    }
}

/// State of a checkbox or radio button control (MS-XLS 2.5.141 `fChecked`).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum CheckState {
    Unchecked,
    Checked,
    Mixed,
}

impl CheckState {
    pub(super) fn from_code(value: u16) -> Option<Self> {
        Some(match value {
            0x0000 => Self::Unchecked,
            0x0001 => Self::Checked,
            0x0002 => Self::Mixed,
            _ => return None,
        })
    }
    pub(super) fn code(self) -> u16 {
        match self {
            Self::Unchecked => 0x0000,
            Self::Checked => 0x0001,
            Self::Mixed => 0x0002,
        }
    }
}

/// Input data validation expected by an edit box (MS-XLS 2.5.144 `ivtEdit`).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum EditBoxValidation {
    AnyString,
    Integer,
    Number,
    Reference,
    Formula,
}

impl EditBoxValidation {
    pub(super) fn from_code(value: u16) -> Option<Self> {
        Some(match value {
            0x0000 => Self::AnyString,
            0x0001 => Self::Integer,
            0x0002 => Self::Number,
            0x0003 => Self::Reference,
            0x0004 => Self::Formula,
            _ => return None,
        })
    }
    pub(super) fn code(self) -> u16 {
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
pub enum ListSelectionType {
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
pub enum ListBehaviorClass {
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
pub enum DropDownStyle {
    Combo,
    ComboEdit,
    Simple,
    /// Value 3; MS-XLS does not define a meaning for it.
    Reserved,
}

/// FtCblsData (MS-XLS 2.5.141): checkbox or radio button properties.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct FtCblsData {
    pub state: CheckState,
    /// Unicode character of the accelerator key; 0 means none.
    pub accelerator: u16,
    pub reserved: u16,
    /// Raw `fNo3d`/`unused` bitfield; the unused bits are preserved verbatim.
    pub flags: u16,
}

impl FtCblsData {
    /// Whether the control is drawn without three-dimensional effects.
    pub fn no_3d(&self) -> bool {
        self.flags & NO_3D != 0
    }
}

/// FtGboData (MS-XLS 2.5.145): group box properties.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct FtGboData {
    /// Unicode character of the accelerator key; 0 means none.
    pub accelerator: u16,
    pub reserved: u16,
    /// Raw `fNo3d`/`unused2` bitfield; the unused bits are preserved verbatim.
    pub flags: u16,
}

impl FtGboData {
    /// Whether the control is drawn without three-dimensional effects.
    pub fn no_3d(&self) -> bool {
        self.flags & NO_3D != 0
    }
}

/// FtEdoData (MS-XLS 2.5.144): edit box properties.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct FtEdoData {
    pub validation: EditBoxValidation,
    pub multi_line: bool,
    pub vertical_scroll_bar: bool,
    /// Identifier of the associated list control; 0 means none.
    pub list_control_id: u16,
}

/// FtRboData (MS-XLS 2.5.153): radio button grouping.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct FtRboData {
    /// Identifier of the next radio button in the group; 0 means none.
    pub next_radio_button_id: u16,
    /// Whether this is the first radio button of its group.
    pub first_in_group: bool,
}

/// FtSbs (MS-XLS 2.5.154): scroll bar or spin control properties.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct FtSbs {
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

impl FtSbs {
    /// Validate the range and non-negative increments required by MS-XLS
    /// 2.5.154. Parsed records that fail this check remain lossless unknown
    /// subrecords; authored typed values fail before serialization.
    pub fn validate(&self) -> Result<()> {
        if self.minimum > self.maximum {
            return Err(invalid(FT_SBS, "FtSbs minimum exceeds maximum"));
        }
        if !(self.minimum..=self.maximum).contains(&self.value) {
            return Err(invalid(FT_SBS, "FtSbs value is outside its range"));
        }
        if self.increment < 0 || self.page_increment < 0 || self.scroll_width < 0 {
            return Err(invalid(
                FT_SBS,
                "FtSbs increments and scroll width must be non-negative",
            ));
        }
        Ok(())
    }

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
pub struct LbsItem {
    pub(super) text: String,
    pub(super) encoded: Vec<u8>,
}

impl LbsItem {
    pub(super) fn parse(encoded: Vec<u8>) -> Option<Self> {
        if u16_at(&encoded, 0)? > 0x00FF {
            return None;
        }
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
            if count > 0x00FF {
                return None;
            }
            encoded.extend_from_slice(&count.to_le_bytes());
            encoded.push(0);
            encoded.extend(text.chars().map(|value| value as u8));
        } else {
            let units = text.encode_utf16().collect::<Vec<_>>();
            let count = u16::try_from(units.len()).ok()?;
            if count > 0x00FF {
                return None;
            }
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
pub struct LbsDropData {
    /// Raw `wStyle`/`fFiltered` bitfield; the unused bits are preserved verbatim.
    pub flags: u16,
    /// Number of lines displayed in the dropdown (`cLine`).
    pub line_count: u16,
    /// Smallest width in pixels allowed for the dropdown window (`dxMin`).
    pub min_width: u16,
    /// Current string value of the dropdown (`str`).
    pub(super) text: LbsItem,
    /// Trailing undefined byte, present iff `str` occupies an odd number of
    /// bytes (`unused3`); preserved verbatim.
    pub padding: Option<u8>,
}

impl LbsDropData {
    /// Validate the bounded dropdown dimensions required by MS-XLS 2.5.171.
    pub fn validate(&self) -> Result<()> {
        if self.line_count > 0x7FFF || self.min_width > 0x7FFF {
            return Err(invalid(
                FT_LBS_DATA,
                "LbsDropData dimensions exceed the MS-XLS limit",
            ));
        }
        Ok(())
    }

    /// The dropdown's visual style (`wStyle`).
    pub fn style(&self) -> DropDownStyle {
        match self.flags & DROP_STYLE_MASK {
            0 => DropDownStyle::Combo,
            1 => DropDownStyle::ComboEdit,
            2 => DropDownStyle::Simple,
            _ => DropDownStyle::Reserved,
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
pub struct FtLbsData {
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
    pub drop_down: Option<LbsDropData>,
    /// List item strings (`rgLines`), each in its original encoding.
    pub(super) items: Vec<LbsItem>,
    /// Per-item multiple-selection state (`bsels`).
    pub multi_selection: Vec<bool>,
    /// Bytes following the typed portion, preserved verbatim on write.
    pub trailing: Vec<u8>,
}

impl FtLbsData {
    /// Validate the list header and any dropdown dimensions required by
    /// MS-XLS 2.5.147 and 2.5.171. Item and selection arrays may be partial
    /// because the owning Obj can continue into later Continue records.
    pub fn validate(&self) -> Result<()> {
        if self.entry_count > 0x7FFF {
            return Err(invalid(FT_LBS_DATA, "FtLbsData entry count exceeds 0x7FFF"));
        }
        if self.selected_index > self.entry_count {
            return Err(invalid(
                FT_LBS_DATA,
                "FtLbsData selected index exceeds entry count",
            ));
        }
        if let Some(drop_down) = &self.drop_down {
            drop_down.validate()?;
        }
        Ok(())
    }

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
    pub fn selection_type(&self) -> ListSelectionType {
        match (self.flags >> LBS_SELECTION_TYPE_SHIFT) & LBS_SELECTION_TYPE_MASK {
            0 => ListSelectionType::Single,
            1 => ListSelectionType::Multi,
            2 => ListSelectionType::CtrlMulti,
            _ => ListSelectionType::Reserved,
        }
    }
    /// The behavior class of the list (`lct`).
    pub fn behavior_class(&self) -> ListBehaviorClass {
        match (self.flags >> LBS_BEHAVIOR_CLASS_SHIFT) as u8 {
            0x00 => ListBehaviorClass::Regular,
            0x01 => ListBehaviorClass::PivotPageField,
            0x03 => ListBehaviorClass::AutoFilter,
            0x05 => ListBehaviorClass::AutoComplete,
            0x06 => ListBehaviorClass::DataValidation,
            0x07 => ListBehaviorClass::PivotField,
            0x09 => ListBehaviorClass::TotalRow,
            value => ListBehaviorClass::Unknown(value),
        }
    }
    /// The list item strings.
    pub fn items(&self) -> &[LbsItem] {
        &self.items
    }
    /// Replace the list item strings.
    pub fn set_items(&mut self, items: Vec<LbsItem>) {
        self.items = items;
    }

    pub(super) fn is_vacant(&self) -> bool {
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
    Unknown { kind: u16, data: Vec<u8> },
    End,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct OleObjectRecord {
    pub subrecords: Vec<ObjSubrecord>,
    /// Complete adjacent TxO record, retained byte-for-byte.
    pub text_object: Option<Vec<u8>>,
}

impl OleObjectRecord {
    pub fn parse(data: &[u8], text_object: Option<Vec<u8>>) -> Result<Self> {
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
                ObjSubrecord::Common(value) => Some(value.object_id),
                _ => None,
            })
            .unwrap_or(0)
    }

    pub fn storage_position(&self) -> Option<u32> {
        self.subrecords.iter().find_map(|value| match value {
            ObjSubrecord::PictureFormula(value) => value.storage_position,
            _ => None,
        })
    }

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

    pub fn validate(&self) -> Result<()> {
        if self.subrecords.len() > 1_024 {
            return Err(invalid(OBJ, "too many Obj subrecords"));
        }
        let common = self
            .subrecords
            .iter()
            .filter_map(|value| match value {
                ObjSubrecord::Common(value) => Some(value),
                _ => None,
            })
            .collect::<Vec<_>>();
        if common.len() != 1
            || common[0].object_type != 8
            || common[0].object_id == 0
            || !matches!(self.subrecords.first(), Some(ObjSubrecord::Common(_)))
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
                ObjSubrecord::PictureFlags(value) => Some(*value),
                _ => None,
            })
            .collect::<Vec<_>>();
        if pio.len() != 1 {
            return Err(invalid(OBJ, "OLE Obj requires one FtPioGrbit"));
        }
        pio[0].validate()?;
        if pio[0].is_control() || pio[0].uses_control_stream() {
            return Err(invalid(
                OBJ,
                "OLE Obj data must be in an embedding or link storage",
            ));
        }
        if self
            .subrecords
            .iter()
            .filter(|value| matches!(value, ObjSubrecord::PictureFormula(_)))
            .count()
            > 1
        {
            return Err(invalid(OBJ, "duplicate FtPictFmla"));
        }
        if !matches!(self.subrecords.last(), Some(ObjSubrecord::End)) {
            return Err(invalid(OBJ, "OLE Obj must end with FtEnd"));
        }
        Ok(())
    }

    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        self.validate()?;
        record(OBJ, &serialize_subrecords(&self.subrecords)?)
    }
}

/// A worksheet form control (checkbox, radio button, edit box, group box,
/// spin/scroll bar, list box, or dropdown) backed by an Obj record.
///
/// Unlike [`OleObjectRecord`], which enforces the strict OLE-object shape,
/// this view accepts any Obj whose `cmo.ot` names a form control and types its
/// control-specific subrecords per MS-XLS 2.5.141-2.5.154.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct FormControl {
    /// Complete Obj subrecord list; the first entry is always
    /// [`ObjSubrecord::Common`].
    pub subrecords: Vec<ObjSubrecord>,
    /// Complete adjacent TxO record, retained byte-for-byte.
    pub text_object: Option<Vec<u8>>,
}

impl FormControl {
    /// Parse a form-control Obj record body. Returns `None` when the record is
    /// not a worksheet form control (picture, chart, note, ...) or its framing
    /// is broken; such records keep flowing through untouched.
    pub fn parse(data: &[u8], text_object: Option<Vec<u8>>) -> Option<Self> {
        let subrecords = parse_subrecords(data).ok()?;
        let common = match subrecords.first() {
            Some(ObjSubrecord::Common(value)) => value,
            _ => return None,
        };
        let control = common
            .object_kind()
            .is_some_and(ObjectType::is_form_control);
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
    pub fn control_type(&self) -> Option<ObjectType> {
        self.common().and_then(FtCmo::object_kind)
    }

    /// Checkbox or radio button state, when present.
    pub fn check_box_data(&self) -> Option<&FtCblsData> {
        self.find(|value| match value {
            ObjSubrecord::CheckBoxData(value) => Some(value),
            _ => None,
        })
    }

    /// Radio button grouping, when present.
    pub fn radio_button_data(&self) -> Option<&FtRboData> {
        self.find(|value| match value {
            ObjSubrecord::RadioButtonData(value) => Some(value),
            _ => None,
        })
    }

    /// Edit box properties, when present.
    pub fn edit_box_data(&self) -> Option<&FtEdoData> {
        self.find(|value| match value {
            ObjSubrecord::EditBoxData(value) => Some(value),
            _ => None,
        })
    }

    /// Group box properties, when present.
    pub fn group_box_data(&self) -> Option<&FtGboData> {
        self.find(|value| match value {
            ObjSubrecord::GroupBoxData(value) => Some(value),
            _ => None,
        })
    }

    /// Scroll bar or spin control properties, when present.
    pub fn scroll_bar_data(&self) -> Option<&FtSbs> {
        self.find(|value| match value {
            ObjSubrecord::ScrollBarData(value) => Some(value),
            _ => None,
        })
    }

    /// List box or dropdown properties, when present.
    pub fn list_box_data(&self) -> Option<&FtLbsData> {
        self.find(|value| match value {
            ObjSubrecord::ListBoxData(value) => Some(value),
            _ => None,
        })
    }

    /// Serialize the control back to a complete OBJ record.
    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        record(OBJ, &serialize_subrecords(&self.subrecords)?)
    }

    fn common(&self) -> Option<&FtCmo> {
        self.find(|value| match value {
            ObjSubrecord::Common(value) => Some(value),
            _ => None,
        })
    }

    fn find<'a, T>(&'a self, pick: impl Fn(&'a ObjSubrecord) -> Option<&'a T>) -> Option<&'a T> {
        self.subrecords.iter().find_map(pick)
    }
}
