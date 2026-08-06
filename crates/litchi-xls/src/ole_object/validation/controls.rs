//! Form-control bounds and semantic invariants.

use super::super::semantic::{FtLbsData, FtSbs, LbsDropData};
use super::super::{FT_LBS_DATA, FT_SBS, invalid};
use crate::error::Result;

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
