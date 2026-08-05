use allsorts::{
    binary::read::ReadScope,
    font_data::FontData as AllsortsFontData,
    subset::{CmapTarget, SubsetProfile, subset},
};

use super::Subsetter;
use crate::model::{FontData, FontError};

/// OpenType subsetter backed by allsorts.
pub struct Allsorts;

impl Default for Allsorts {
    fn default() -> Self {
        Self
    }
}

impl Allsorts {
    pub fn new() -> Self {
        Self
    }
}

impl Subsetter for Allsorts {
    fn subset(&self, font: &FontData, glyph_ids: &[u16]) -> Result<Vec<u8>, FontError> {
        let scope = ReadScope::new(&font.data);
        let font_data = scope
            .read::<AllsortsFontData<'_>>()
            .map_err(|e| FontError::SubsettingFailed(e.to_string()))?;

        let provider = font_data
            .table_provider(font.index as usize)
            .map_err(|e| FontError::SubsettingFailed(e.to_string()))?;

        // Basic subsetting with PDF profile and default CmapTarget
        let subset_font = subset(
            &provider,
            glyph_ids,
            &SubsetProfile::Pdf,
            CmapTarget::default(),
        )
        .map_err(|e| FontError::SubsettingFailed(e.to_string()))?;

        Ok(subset_font)
    }
}
