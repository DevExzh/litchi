//! Semantic metadata for worksheet form controls.

/// State of a checkbox or radio button control (`fChecked`, MS-XLS 2.5.141).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum CheckState {
    Unchecked,
    Checked,
    Mixed,
}

impl CheckState {
    pub(crate) fn from_code(value: u16) -> Option<Self> {
        Some(match value {
            0x0000 => Self::Unchecked,
            0x0001 => Self::Checked,
            0x0002 => Self::Mixed,
            _ => return None,
        })
    }

    pub(crate) fn code(self) -> u16 {
        match self {
            Self::Unchecked => 0x0000,
            Self::Checked => 0x0001,
            Self::Mixed => 0x0002,
        }
    }
}

/// Input validation expected by an edit box (`ivtEdit`, MS-XLS 2.5.144).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum EditBoxValidation {
    AnyString,
    Integer,
    Number,
    Reference,
    Formula,
}

impl EditBoxValidation {
    pub(crate) fn from_code(value: u16) -> Option<Self> {
        Some(match value {
            0x0000 => Self::AnyString,
            0x0001 => Self::Integer,
            0x0002 => Self::Number,
            0x0003 => Self::Reference,
            0x0004 => Self::Formula,
            _ => return None,
        })
    }

    pub(crate) fn code(self) -> u16 {
        match self {
            Self::AnyString => 0x0000,
            Self::Integer => 0x0001,
            Self::Number => 0x0002,
            Self::Reference => 0x0003,
            Self::Formula => 0x0004,
        }
    }
}

/// Selection behavior of a list control (`wListSelType`, MS-XLS 2.5.147).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ListSelectionType {
    Single,
    Multi,
    CtrlMulti,
    Reserved,
}

/// Behavior class of a list control (`lct`, MS-XLS 2.5.147).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ListBehaviorClass {
    Regular,
    PivotPageField,
    AutoFilter,
    AutoComplete,
    DataValidation,
    PivotField,
    TotalRow,
    Unknown(u8),
}

/// Visual style of a dropdown control (`wStyle`, MS-XLS 2.5.171).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum DropDownStyle {
    Combo,
    ComboEdit,
    Simple,
    Reserved,
}

/// Checkbox or radio-button properties (`FtCblsData`, MS-XLS 2.5.141).
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct FtCblsData {
    pub state: CheckState,
    pub accelerator: u16,
    pub reserved: u16,
    /// Raw `fNo3d`/unused bitfield, including undefined bits.
    pub flags: u16,
}

impl FtCblsData {
    pub fn no_3d(&self) -> bool {
        self.flags & super::super::NO_3D != 0
    }
}

/// Group-box properties (`FtGboData`, MS-XLS 2.5.145).
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct FtGboData {
    pub accelerator: u16,
    pub reserved: u16,
    /// Raw `fNo3d`/unused bitfield, including undefined bits.
    pub flags: u16,
}

impl FtGboData {
    pub fn no_3d(&self) -> bool {
        self.flags & super::super::NO_3D != 0
    }
}

/// Edit-box properties (`FtEdoData`, MS-XLS 2.5.144).
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct FtEdoData {
    pub validation: EditBoxValidation,
    pub multi_line: bool,
    pub vertical_scroll_bar: bool,
    pub list_control_id: u16,
}

/// Radio-button grouping (`FtRboData`, MS-XLS 2.5.153).
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct FtRboData {
    pub next_radio_button_id: u16,
    pub first_in_group: bool,
}

/// Scroll-bar or spin-control properties (`FtSbs`, MS-XLS 2.5.154).
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct FtSbs {
    pub reserved: [u8; 4],
    pub value: i16,
    pub minimum: i16,
    pub maximum: i16,
    pub increment: i16,
    pub page_increment: i16,
    pub horizontal: bool,
    pub scroll_width: i16,
    /// Raw fDraw/fDrawSliderOnly/fTrackElevator/fNo3d bitfield.
    pub flags: u16,
}

impl FtSbs {
    pub fn draw(&self) -> bool {
        self.flags & super::super::SBS_DRAW != 0
    }

    pub fn draw_slider_only(&self) -> bool {
        self.flags & super::super::SBS_DRAW_SLIDER_ONLY != 0
    }

    pub fn track_elevator(&self) -> bool {
        self.flags & super::super::SBS_TRACK_ELEVATOR != 0
    }

    pub fn no_3d(&self) -> bool {
        self.flags & super::super::SBS_NO_3D != 0
    }
}

/// A list item retaining its original XLUnicodeString encoding.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct LbsItem {
    pub(crate) text: String,
    pub(crate) encoded: Vec<u8>,
}

impl LbsItem {
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
            encoded.push(super::super::XL_STRING_HIGH_BYTE);
            encoded.extend(units.iter().flat_map(|unit| unit.to_le_bytes()));
        }
        Some(Self {
            text: text.to_string(),
            encoded,
        })
    }

    pub fn text(&self) -> &str {
        &self.text
    }

    pub fn encoded(&self) -> &[u8] {
        &self.encoded
    }
}

/// Dropdown-specific list-box properties (`LbsDropData`, MS-XLS 2.5.171).
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct LbsDropData {
    pub flags: u16,
    pub line_count: u16,
    pub min_width: u16,
    pub(crate) text: LbsItem,
    pub padding: Option<u8>,
}

impl LbsDropData {
    pub fn style(&self) -> DropDownStyle {
        match self.flags & super::super::DROP_STYLE_MASK {
            0 => DropDownStyle::Combo,
            1 => DropDownStyle::ComboEdit,
            2 => DropDownStyle::Simple,
            _ => DropDownStyle::Reserved,
        }
    }

    pub fn filtered(&self) -> bool {
        self.flags & super::super::DROP_FILTERED != 0
    }

    pub fn text(&self) -> &str {
        self.text.text()
    }
}

/// List-box or dropdown properties (`FtLbsData`, MS-XLS 2.5.147).
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct FtLbsData {
    pub formula: Vec<u8>,
    pub entry_count: u16,
    pub selected_index: u16,
    pub flags: u16,
    pub edit_box_id: u16,
    pub drop_down: Option<LbsDropData>,
    pub(crate) items: Vec<LbsItem>,
    pub multi_selection: Vec<bool>,
    pub trailing: Vec<u8>,
}

impl FtLbsData {
    pub fn has_behavior_class(&self) -> bool {
        self.flags & super::super::LBS_USE_CB != 0
    }

    pub fn has_item_strings(&self) -> bool {
        self.flags & super::super::LBS_VALID_PLEX != 0
    }

    pub fn has_edit_box(&self) -> bool {
        self.flags & super::super::LBS_VALID_IDS != 0
    }

    pub fn no_3d(&self) -> bool {
        self.flags & super::super::LBS_NO_3D != 0
    }

    pub fn selection_type(&self) -> ListSelectionType {
        match (self.flags >> super::super::LBS_SELECTION_TYPE_SHIFT)
            & super::super::LBS_SELECTION_TYPE_MASK
        {
            0 => ListSelectionType::Single,
            1 => ListSelectionType::Multi,
            2 => ListSelectionType::CtrlMulti,
            _ => ListSelectionType::Reserved,
        }
    }

    pub fn behavior_class(&self) -> ListBehaviorClass {
        match (self.flags >> super::super::LBS_BEHAVIOR_CLASS_SHIFT) as u8 {
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

    pub fn items(&self) -> &[LbsItem] {
        &self.items
    }

    pub fn set_items(&mut self, items: Vec<LbsItem>) {
        self.items = items;
    }

    pub(crate) fn is_vacant(&self) -> bool {
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

/// A worksheet form control backed by an Obj record.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct FormControl {
    pub subrecords: Vec<super::obj::ObjSubrecord>,
    pub text_object: Option<Vec<u8>>,
}

impl FormControl {
    pub fn object_id(&self) -> u16 {
        self.common().map_or(0, |value| value.object_id)
    }

    pub fn control_type(&self) -> Option<super::obj::ObjectType> {
        self.common().and_then(super::obj::FtCmo::object_kind)
    }

    pub fn check_box_data(&self) -> Option<&FtCblsData> {
        self.find(|value| match value {
            super::obj::ObjSubrecord::CheckBoxData(value) => Some(value),
            _ => None,
        })
    }

    pub fn radio_button_data(&self) -> Option<&FtRboData> {
        self.find(|value| match value {
            super::obj::ObjSubrecord::RadioButtonData(value) => Some(value),
            _ => None,
        })
    }

    pub fn edit_box_data(&self) -> Option<&FtEdoData> {
        self.find(|value| match value {
            super::obj::ObjSubrecord::EditBoxData(value) => Some(value),
            _ => None,
        })
    }

    pub fn group_box_data(&self) -> Option<&FtGboData> {
        self.find(|value| match value {
            super::obj::ObjSubrecord::GroupBoxData(value) => Some(value),
            _ => None,
        })
    }

    pub fn scroll_bar_data(&self) -> Option<&FtSbs> {
        self.find(|value| match value {
            super::obj::ObjSubrecord::ScrollBarData(value) => Some(value),
            _ => None,
        })
    }

    pub fn list_box_data(&self) -> Option<&FtLbsData> {
        self.find(|value| match value {
            super::obj::ObjSubrecord::ListBoxData(value) => Some(value),
            _ => None,
        })
    }

    fn common(&self) -> Option<&super::obj::FtCmo> {
        self.find(|value| match value {
            super::obj::ObjSubrecord::Common(value) => Some(value),
            _ => None,
        })
    }

    fn find<'a, T>(
        &'a self,
        pick: impl Fn(&'a super::obj::ObjSubrecord) -> Option<&'a T>,
    ) -> Option<&'a T> {
        self.subrecords.iter().find_map(pick)
    }
}
