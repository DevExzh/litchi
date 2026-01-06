use allsorts::{
    binary::read::ReadScope,
    font_data::FontData as AllsortsFontData,
    subset::{CmapTarget, SubsetProfile, subset},
};

use crate::fonts::{FontData, FontError, FontSubsetter};

pub struct AllsortsSubsetter;

impl Default for AllsortsSubsetter {
    fn default() -> Self {
        Self
    }
}

impl AllsortsSubsetter {
    pub fn new() -> Self {
        Self
    }
}

impl FontSubsetter for AllsortsSubsetter {
    fn subset(&self, font: &FontData, glyph_ids: &[u16]) -> Result<Vec<u8>, FontError> {
        let scope = ReadScope::new(&font.data);
        let font_data = scope
            .read::<AllsortsFontData>()
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
