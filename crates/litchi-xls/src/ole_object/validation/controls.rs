//! Form-control bounds and semantic invariants.

use super::super::semantic::{FormControl, FtLbsData, FtSbs, LbsDropData, ObjSubrecord};
use super::super::{FT_LBS_DATA, FT_SBS, OBJ, invalid};
use crate::error::Result;

impl FormControl {
    /// Validates the structural and type-specific invariants required when a
    /// new form-control Obj is authored.
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
        if !matches!(self.subrecords.last(), Some(ObjSubrecord::End)) {
            return Err(invalid(OBJ, "form-control Obj must end with FtEnd"));
        }

        for value in &self.subrecords {
            match value {
                ObjSubrecord::ScrollBarData(value) => value.validate()?,
                ObjSubrecord::ListBoxData(value) => value.validate()?,
                _ => {},
            }
        }

        match object_type {
            super::super::semantic::ObjectType::CheckBox => {
                if self.check_box_data().is_none() {
                    return Err(invalid(OBJ, "checkbox Obj requires FtCblsData"));
                }
            },
            super::super::semantic::ObjectType::RadioButton => {
                let Some(data) = self.check_box_data() else {
                    return Err(invalid(OBJ, "radio-button Obj requires FtCblsData"));
                };
                if data.state == super::super::semantic::CheckState::Mixed {
                    return Err(invalid(
                        OBJ,
                        "radio-button Obj cannot use the mixed checkbox state",
                    ));
                }
                if self.radio_button_data().is_none() {
                    return Err(invalid(OBJ, "radio-button Obj requires FtRboData"));
                }
            },
            super::super::semantic::ObjectType::EditBox => {
                if self.edit_box_data().is_none() {
                    return Err(invalid(OBJ, "edit-box Obj requires FtEdoData"));
                }
            },
            super::super::semantic::ObjectType::SpinControl
            | super::super::semantic::ObjectType::ScrollBar => {
                if self.scroll_bar_data().is_none() {
                    return Err(invalid(OBJ, "scroll/spin Obj requires FtSbs"));
                }
            },
            super::super::semantic::ObjectType::List
            | super::super::semantic::ObjectType::DropDown => {
                if self.list_box_data().is_none() {
                    return Err(invalid(OBJ, "list/dropdown Obj requires FtLbsData"));
                }
            },
            super::super::semantic::ObjectType::GroupBox => {
                if self.group_box_data().is_none() {
                    return Err(invalid(OBJ, "group-box Obj requires FtGboData"));
                }
            },
            super::super::semantic::ObjectType::Label
            | super::super::semantic::ObjectType::DialogBox => {
                return Err(invalid(
                    OBJ,
                    "label/dialog form-control authoring is not modeled",
                ));
            },
            _ => return Err(invalid(OBJ, "unsupported worksheet form-control type")),
        }
        Ok(())
    }
}

impl FtSbs {
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
    pub fn validate(&self) -> Result<()> {
        if self.line_count > 0x7FFF || self.min_width > 0x7FFF {
            return Err(invalid(
                FT_LBS_DATA,
                "LbsDropData dimensions exceed the MS-XLS limit",
            ));
        }
        Ok(())
    }
}

impl FtLbsData {
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
}
