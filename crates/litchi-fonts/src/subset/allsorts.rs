use ::allsorts::{
    binary::read::ReadScope,
    font_data::FontData as AllsortsFontData,
    subset::{CmapTarget, SubsetProfile, subset},
};

use super::{Limits, SubsetSession, Subsetter, parse_error, read_write_error, subset_error};
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
    fn subset(
        &self,
        font: &FontData,
        glyph_ids: &[u16],
        limits: Limits,
    ) -> Result<Vec<u8>, FontError> {
        let session = SubsetSession::new(limits);
        session.charge_subset_staging(font.data.len(), glyph_ids.len())?;
        self.subset_scoped(&session, font, glyph_ids)
    }
}

impl Allsorts {
    pub(crate) fn subset_scoped(
        &self,
        _session: &SubsetSession,
        font: &FontData,
        glyph_ids: &[u16],
    ) -> Result<Vec<u8>, FontError> {
        let scope = ReadScope::new(&font.data);
        let font_data = scope
            .read::<AllsortsFontData<'_>>()
            .map_err(|error| parse_error("failed to parse font", error))?;

        let provider = font_data
            .table_provider(font.index as usize)
            .map_err(|error| read_write_error("failed to get font face", error))?;

        // Basic subsetting with PDF profile and default CmapTarget
        let subset_font = subset(
            &provider,
            glyph_ids,
            &SubsetProfile::Pdf,
            CmapTarget::default(),
        )
        .map_err(subset_error)?;

        Ok(subset_font)
    }
}
