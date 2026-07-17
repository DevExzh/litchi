//! Typed character baseline, expansion, scaling, and kerning state.

use crate::{RtfError, RtfResult};

pub const MAX_CHARACTER_BASELINE_HALF_POINTS: i32 = 31_680;
pub const MAX_CHARACTER_EXPANSION: i32 = 31_680;
pub const MAX_CHARACTER_SCALE_PERCENT: i32 = 600;
pub const MAX_CHARACTER_KERNING_HALF_POINTS: i32 = 32_767;

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum CharacterBaseline {
    #[default]
    Normal,
    Superscript,
    Subscript,
    RaisedHalfPoints(u16),
    LoweredHalfPoints(u16),
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum CharacterExpansion {
    #[default]
    None,
    QuarterPoints(i16),
    Twips(i16),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CharacterPositioning {
    pub baseline: CharacterBaseline,
    pub expansion: CharacterExpansion,
    pub horizontal_scale_percent: u16,
    pub kerning_half_points: u16,
}

impl Default for CharacterPositioning {
    fn default() -> Self {
        Self {
            baseline: CharacterBaseline::Normal,
            expansion: CharacterExpansion::None,
            horizontal_scale_percent: 100,
            kerning_half_points: 0,
        }
    }
}

impl CharacterPositioning {
    fn bounded_nonnegative(value: i32, maximum: i32, name: &str) -> RtfResult<u16> {
        if !(0..=maximum).contains(&value) {
            return Err(RtfError::MalformedDocument(format!("RTF {name} parameter is out of range")));
        }
        Ok(value as u16)
    }

    pub(crate) fn set_superscript(&mut self, enabled: bool) {
        if enabled {
            self.baseline = CharacterBaseline::Superscript;
        } else if self.baseline == CharacterBaseline::Superscript {
            self.baseline = CharacterBaseline::Normal;
        }
    }

    pub(crate) fn set_subscript(&mut self, enabled: bool) {
        if enabled {
            self.baseline = CharacterBaseline::Subscript;
        } else if self.baseline == CharacterBaseline::Subscript {
            self.baseline = CharacterBaseline::Normal;
        }
    }

    pub(crate) fn clear_baseline(&mut self) {
        self.baseline = CharacterBaseline::Normal;
    }

    pub(crate) fn set_raised(&mut self, value: i32) -> RtfResult<()> {
        let value = Self::bounded_nonnegative(value, MAX_CHARACTER_BASELINE_HALF_POINTS, "up")?;
        self.baseline = if value == 0 { CharacterBaseline::Normal } else { CharacterBaseline::RaisedHalfPoints(value) };
        Ok(())
    }

    pub(crate) fn set_lowered(&mut self, value: i32) -> RtfResult<()> {
        let value = Self::bounded_nonnegative(value, MAX_CHARACTER_BASELINE_HALF_POINTS, "dn")?;
        self.baseline = if value == 0 { CharacterBaseline::Normal } else { CharacterBaseline::LoweredHalfPoints(value) };
        Ok(())
    }

    pub(crate) fn set_quarter_point_expansion(&mut self, value: i32) -> RtfResult<()> {
        if !(-MAX_CHARACTER_EXPANSION..=MAX_CHARACTER_EXPANSION).contains(&value) {
            return Err(RtfError::MalformedDocument("RTF expnd parameter is out of range".to_string()));
        }
        self.expansion = if value == 0 { CharacterExpansion::None } else { CharacterExpansion::QuarterPoints(value as i16) };
        Ok(())
    }

    pub(crate) fn set_twip_expansion(&mut self, value: i32) -> RtfResult<()> {
        if !(-MAX_CHARACTER_EXPANSION..=MAX_CHARACTER_EXPANSION).contains(&value) {
            return Err(RtfError::MalformedDocument("RTF expndtw parameter is out of range".to_string()));
        }
        self.expansion = if value == 0 { CharacterExpansion::None } else { CharacterExpansion::Twips(value as i16) };
        Ok(())
    }

    pub(crate) fn set_scale(&mut self, value: i32) -> RtfResult<()> {
        if !(1..=MAX_CHARACTER_SCALE_PERCENT).contains(&value) {
            return Err(RtfError::MalformedDocument("RTF charscalex parameter is out of range".to_string()));
        }
        self.horizontal_scale_percent = value as u16;
        Ok(())
    }

    pub(crate) fn set_kerning(&mut self, value: i32) -> RtfResult<()> {
        self.kerning_half_points = Self::bounded_nonnegative(value, MAX_CHARACTER_KERNING_HALF_POINTS, "kerning")?;
        Ok(())
    }

    pub fn validate(&self) -> RtfResult<()> {
        match self.baseline {
            CharacterBaseline::RaisedHalfPoints(value) | CharacterBaseline::LoweredHalfPoints(value)
                if i32::from(value) > MAX_CHARACTER_BASELINE_HALF_POINTS => return Err(RtfError::MalformedDocument("RTF character baseline is out of range".to_string())),
            _ => {}
        }
        match self.expansion {
            CharacterExpansion::QuarterPoints(value) | CharacterExpansion::Twips(value)
                if i32::from(value).unsigned_abs() > MAX_CHARACTER_EXPANSION as u32 => return Err(RtfError::MalformedDocument("RTF character expansion is out of range".to_string())),
            _ => {}
        }
        if !(1..=MAX_CHARACTER_SCALE_PERCENT as u16).contains(&self.horizontal_scale_percent) {
            return Err(RtfError::MalformedDocument("RTF character scale is out of range".to_string()));
        }
        if i32::from(self.kerning_half_points) > MAX_CHARACTER_KERNING_HALF_POINTS {
            return Err(RtfError::MalformedDocument("RTF character kerning is out of range".to_string()));
        }
        Ok(())
    }
}
