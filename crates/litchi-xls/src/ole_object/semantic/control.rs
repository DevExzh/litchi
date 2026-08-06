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

impl DropDownStyle {
    pub const fn code(self) -> u16 {
        match self {
            Self::Combo => 0,
            Self::ComboEdit => 1,
            Self::Simple => 2,
            Self::Reserved => 3,
        }
    }
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

    /// Creates an item or reports the MS-XLS XLUnicodeString size limit.
    pub fn try_new(text: &str) -> crate::error::Result<Self> {
        Self::new(text).ok_or_else(|| {
            super::super::invalid(
                super::super::FT_LBS_DATA,
                "list item exceeds the MS-XLS XLUnicodeString limit",
            )
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
    /// Creates dropdown metadata from a typed style and display string.
    pub fn new(style: DropDownStyle, line_count: u16, min_width: u16, text: &str) -> Option<Self> {
        Self::try_new(style, line_count, min_width, text).ok()
    }

    /// Fallible form of [`LbsDropData::new`], with MS-XLS validation errors.
    pub fn try_new(
        style: DropDownStyle,
        line_count: u16,
        min_width: u16,
        text: &str,
    ) -> crate::error::Result<Self> {
        let item = LbsItem::try_new(text)?;
        Self::from_item(style, line_count, min_width, item).ok_or_else(|| {
            super::super::invalid(
                super::super::FT_LBS_DATA,
                "dropdown metadata uses a reserved style or exceeds its dimensions",
            )
        })
    }

    /// Creates dropdown metadata while retaining a pre-encoded list item.
    pub fn from_item(
        style: DropDownStyle,
        line_count: u16,
        min_width: u16,
        text: LbsItem,
    ) -> Option<Self> {
        if style == DropDownStyle::Reserved || line_count > 0x7FFF || min_width > 0x7FFF {
            return None;
        }
        let padding = (text.encoded.len() % 2 == 1).then_some(0);
        Some(Self {
            flags: style.code(),
            line_count,
            min_width,
            text,
            padding,
        })
    }

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
    /// A source payload ended before all continued list fields were present.
    /// This is deliberately not public: authored values must be complete, but
    /// parsed values retain enough state to round-trip a partial wire payload.
    pub(crate) partial: bool,
}

impl FtLbsData {
    /// Creates complete list metadata from already encoded item values.
    pub fn from_items(items: Vec<LbsItem>) -> crate::error::Result<Self> {
        let entry_count = u16::try_from(items.len()).map_err(|_| {
            super::super::invalid(
                super::super::FT_LBS_DATA,
                "list item count exceeds the u16 representation",
            )
        })?;
        if entry_count > 0x7FFF {
            return Err(super::super::invalid(
                super::super::FT_LBS_DATA,
                "list item count exceeds the MS-XLS limit",
            ));
        }
        Ok(Self {
            entry_count,
            flags: super::super::LBS_VALID_PLEX,
            items,
            ..Self::default()
        })
    }

    /// Creates complete list metadata from display strings.
    pub fn from_texts<I, T>(texts: I) -> crate::error::Result<Self>
    where
        I: IntoIterator<Item = T>,
        T: AsRef<str>,
    {
        let items = texts
            .into_iter()
            .map(|text| LbsItem::try_new(text.as_ref()))
            .collect::<crate::error::Result<Vec<_>>>()?;
        Self::from_items(items)
    }

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

    /// Replaces authored items and synchronizes `cLines` and `fValidPlex`.
    pub fn set_items_checked(&mut self, items: Vec<LbsItem>) -> crate::error::Result<()> {
        let replacement = Self::from_items(items)?;
        self.entry_count = replacement.entry_count;
        self.flags |= super::super::LBS_VALID_PLEX;
        self.items = replacement.items;
        self.partial = false;
        self.trailing.clear();
        if self.selection_type() == ListSelectionType::Single {
            self.multi_selection.clear();
        } else {
            self.multi_selection = vec![false; usize::from(self.entry_count)];
        }
        if self.selected_index > self.entry_count {
            self.selected_index = 0;
        }
        Ok(())
    }

    /// Replaces the item slice without changing other raw fields.
    ///
    /// This is useful while decoding or deliberately preserving a partially
    /// continued source value. New authored values should use
    /// [`FtLbsData::from_items`] or [`FtLbsData::set_items_checked`].
    pub fn set_items(&mut self, items: Vec<LbsItem>) {
        self.items = items;
    }

    pub(crate) fn set_wire_items(&mut self, items: Vec<LbsItem>, partial: bool) {
        self.items = items;
        self.partial = partial;
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
            && !self.partial
    }
}

/// A worksheet form control backed by an Obj record.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct FormControl {
    pub subrecords: Vec<super::obj::ObjSubrecord>,
    pub text_object: Option<Vec<u8>>,
}

impl FormControl {
    /// Creates a form-control Obj from typed payload subrecords.
    ///
    /// The common metadata and terminator are supplied by this constructor;
    /// type-specific requirements are checked by [`FormControl::validate`].
    pub fn new(
        object_type: super::obj::ObjectType,
        object_id: u16,
        payload: impl IntoIterator<Item = super::obj::ObjSubrecord>,
    ) -> Self {
        let mut subrecords = Vec::new();
        subrecords.push(super::obj::ObjSubrecord::Common(super::obj::FtCmo {
            object_type: object_type.code(),
            object_id,
            flags: 0,
            reserved: [0; 12],
        }));
        subrecords.extend(payload);
        if !matches!(subrecords.last(), Some(super::obj::ObjSubrecord::End)) {
            subrecords.push(super::obj::ObjSubrecord::End);
        }
        Self {
            subrecords,
            text_object: None,
        }
    }

    /// Attaches an already framed TxO object to this control.
    pub fn with_text_object(mut self, text_object: Vec<u8>) -> Self {
        self.text_object = Some(text_object);
        self
    }

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
