//! Form-control bounds and semantic invariants.

use super::super::semantic::{
    FormControl, FtLbsData, FtSbs, LbsDropData, ObjSubrecord, ObjectType,
};
use super::super::{FT_LBS_DATA, FT_SBS, OBJ, invalid};
use crate::error::Result;

impl FormControl {
    /// Validates the structural and type-specific invariants required when a
    /// new form-control Obj is authored.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn validate(&self) -> Result<()> {
        if self.subrecords.is_empty() || self.subrecords.len() > 1_024 {
            return Err(invalid(
                OBJ,
                "form-control Obj has an invalid subrecord count",
            ));
        }
        let common = match self.subrecords.first() {
            Some(ObjSubrecord::Common(value)) => value,
            _ => return Err(invalid(OBJ, "form-control Obj requires a leading FtCmo")),
        };
        let object_type = common
            .object_kind()
            .ok_or_else(|| invalid(OBJ, "form-control Obj has an unknown object type"))?;
        if !object_type.is_form_control() {
            return Err(invalid(OBJ, "Obj type is not a worksheet form control"));
        }
        if common.object_id == 0 {
            return Err(invalid(OBJ, "form-control Obj ID must be non-zero"));
        }
        if self
            .subrecords
            .iter()
            .filter(|value| matches!(value, ObjSubrecord::Common(_)))
            .count()
            != 1
        {
            return Err(invalid(OBJ, "form-control Obj requires exactly one FtCmo"));
        }
        if self
            .subrecords
            .iter()
            .filter(|value| matches!(value, ObjSubrecord::End))
            .count()
            != 1
        {
            return Err(invalid(OBJ, "form-control Obj requires exactly one FtEnd"));
        }
        if !matches!(self.subrecords.last(), Some(ObjSubrecord::End)) {
            return Err(invalid(OBJ, "form-control Obj must end with FtEnd"));
        }

        let check_box_count = self
            .subrecords
            .iter()
            .filter(|value| matches!(value, ObjSubrecord::CheckBoxData(_)))
            .count();
        let radio_button_count = self
            .subrecords
            .iter()
            .filter(|value| matches!(value, ObjSubrecord::RadioButtonData(_)))
            .count();
        let edit_box_count = self
            .subrecords
            .iter()
            .filter(|value| matches!(value, ObjSubrecord::EditBoxData(_)))
            .count();
        let group_box_count = self
            .subrecords
            .iter()
            .filter(|value| matches!(value, ObjSubrecord::GroupBoxData(_)))
            .count();
        let scroll_bar_count = self
            .subrecords
            .iter()
            .filter(|value| matches!(value, ObjSubrecord::ScrollBarData(_)))
            .count();
        let list_box_count = self
            .subrecords
            .iter()
            .filter(|value| matches!(value, ObjSubrecord::ListBoxData(_)))
            .count();

        for value in &self.subrecords {
            match value {
                ObjSubrecord::ScrollBarData(value) => value.validate()?,
                ObjSubrecord::ListBoxData(value) => value.validate()?,
                _ => {},
            }
        }

        match object_type {
            ObjectType::CheckBox => {
                if check_box_count != 1 {
                    return Err(invalid(OBJ, "checkbox Obj requires exactly one FtCblsData"));
                }
            },
            ObjectType::RadioButton => {
                if check_box_count != 1 || radio_button_count != 1 {
                    return Err(invalid(
                        OBJ,
                        "radio-button Obj requires one FtCblsData and one FtRboData",
                    ));
                }
                let Some(data) = self.check_box_data() else {
                    return Err(invalid(OBJ, "radio-button Obj requires FtCblsData"));
                };
                if data.state == super::super::semantic::CheckState::Mixed {
                    return Err(invalid(
                        OBJ,
                        "radio-button Obj cannot use the mixed checkbox state",
                    ));
                }
            },
            ObjectType::EditBox => {
                if edit_box_count != 1 {
                    return Err(invalid(OBJ, "edit-box Obj requires exactly one FtEdoData"));
                }
            },
            ObjectType::SpinControl | ObjectType::ScrollBar => {
                if scroll_bar_count != 1 {
                    return Err(invalid(OBJ, "scroll/spin Obj requires exactly one FtSbs"));
                }
            },
            ObjectType::List | ObjectType::DropDown => {
                if scroll_bar_count != 1 || list_box_count != 1 {
                    return Err(invalid(
                        OBJ,
                        "list/dropdown Obj requires exactly one FtSbs and one FtLbsData",
                    ));
                }
                let data = self
                    .list_box_data()
                    .ok_or_else(|| invalid(OBJ, "list/dropdown Obj requires FtLbsData"))?;
                match object_type {
                    ObjectType::List if data.drop_down.is_some() => {
                        return Err(invalid(OBJ, "list Obj cannot contain LbsDropData"));
                    },
                    ObjectType::DropDown if data.drop_down.is_none() => {
                        return Err(invalid(OBJ, "dropdown Obj requires LbsDropData"));
                    },
                    _ => {},
                }
            },
            ObjectType::GroupBox => {
                if group_box_count != 1 {
                    return Err(invalid(OBJ, "group-box Obj requires exactly one FtGboData"));
                }
            },
            ObjectType::Label | ObjectType::DialogBox => {},
            _ => return Err(invalid(OBJ, "unsupported worksheet form-control type")),
        }
        Ok(())
    }
}

impl FtSbs {
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
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
}

impl LbsDropData {
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn validate(&self) -> Result<()> {
        if self.style() == super::super::semantic::DropDownStyle::Reserved {
            return Err(invalid(
                FT_LBS_DATA,
                "LbsDropData uses a reserved dropdown style",
            ));
        }
        if self.line_count > 0x7FFF || self.min_width > 0x7FFF {
            return Err(invalid(
                FT_LBS_DATA,
                "LbsDropData dimensions exceed the MS-XLS limit",
            ));
        }
        if (self.text.encoded.len() % 2 == 1) != self.padding.is_some() {
            return Err(invalid(
                FT_LBS_DATA,
                "LbsDropData padding does not match the XLUnicodeString size",
            ));
        }
        Ok(())
    }
}

impl FtLbsData {
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn validate(&self) -> Result<()> {
        if self.formula.len() > usize::from(u16::MAX) {
            return Err(invalid(FT_LBS_DATA, "FtLbsData formula exceeds u16"));
        }
        if self.entry_count > 0x7FFF {
            return Err(invalid(FT_LBS_DATA, "FtLbsData entry count exceeds 0x7FFF"));
        }
        if self.selected_index > self.entry_count {
            return Err(invalid(
                FT_LBS_DATA,
                "FtLbsData selected index exceeds entry count",
            ));
        }
        if self.items.len() > usize::from(self.entry_count) {
            return Err(invalid(
                FT_LBS_DATA,
                "FtLbsData contains more items than cLines",
            ));
        }
        if self.has_item_strings() {
            if self.items.len() != usize::from(self.entry_count) && !self.partial {
                return Err(invalid(
                    FT_LBS_DATA,
                    "FtLbsData item strings do not satisfy cLines",
                ));
            }
        } else if !self.items.is_empty() {
            return Err(invalid(FT_LBS_DATA, "FtLbsData items require fValidPlex"));
        }
        match self.selection_type() {
            super::super::semantic::ListSelectionType::Single => {
                if !self.multi_selection.is_empty() {
                    return Err(invalid(
                        FT_LBS_DATA,
                        "single-selection FtLbsData cannot contain bsels",
                    ));
                }
            },
            super::super::semantic::ListSelectionType::Multi
            | super::super::semantic::ListSelectionType::CtrlMulti => {
                if self.multi_selection.len() > usize::from(self.entry_count) {
                    return Err(invalid(
                        FT_LBS_DATA,
                        "FtLbsData contains more bsels than cLines",
                    ));
                }
                if self.multi_selection.len() != usize::from(self.entry_count) && !self.partial {
                    return Err(invalid(
                        FT_LBS_DATA,
                        "multiple-selection FtLbsData does not contain one bsel per entry",
                    ));
                }
            },
            super::super::semantic::ListSelectionType::Reserved => {
                return Err(invalid(
                    FT_LBS_DATA,
                    "FtLbsData uses a reserved selection type",
                ));
            },
        }
        if let Some(drop_down) = &self.drop_down {
            drop_down.validate()?;
        }
        Ok(())
    }
}
