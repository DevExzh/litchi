//! Unicode-to-glyph mapping for OpenType faces.

use allsorts::{
    binary::read::ReadScope,
    font::read_cmap_subtable,
    tables::{FontTableProvider, OpenTypeFont, cmap::Cmap},
};

use crate::model::{FontData, FontError, Glyphs};

/// Map a validated Unicode request to sorted, deduplicated glyph IDs.
pub fn glyph_ids(font_data: &FontData, codepoints: &Glyphs) -> Result<Vec<u16>, FontError> {
    let scope = ReadScope::new(&font_data.data);
    let font_file = scope
        .read::<OpenTypeFont<'_>>()
        .map_err(|error| FontError::SubsettingFailed(format!("failed to parse font: {error}")))?;

    let font_index = usize::try_from(font_data.index)
        .map_err(|_| FontError::InvalidFaceIndex(font_data.index))?;
    let provider = font_file.table_provider(font_index).map_err(|error| {
        FontError::SubsettingFailed(format!("failed to get font face: {error}"))
    })?;

    let cmap_data = provider
        .table_data(allsorts::tag::CMAP)
        .map_err(|error| FontError::SubsettingFailed(format!("failed to get cmap table: {error}")))?
        .ok_or_else(|| FontError::SubsettingFailed("font has no cmap table".into()))?;
    let cmap = ReadScope::new(cmap_data.as_ref())
        .read::<Cmap<'_>>()
        .map_err(|error| FontError::SubsettingFailed(format!("failed to read cmap: {error}")))?;
    let subtable = read_cmap_subtable(&cmap)
        .map_err(|error| {
            FontError::SubsettingFailed(format!("failed to read cmap subtable: {error}"))
        })?
        .map(|(_, subtable)| subtable)
        .ok_or_else(|| FontError::SubsettingFailed("no usable cmap subtable found".into()))?;

    let codepoint_count = usize::try_from(codepoints.len()).map_err(|_| {
        FontError::SubsettingFailed("font glyph request exceeds platform limits".into())
    })?;
    let capacity = codepoint_count
        .checked_add(1)
        .ok_or_else(|| FontError::SubsettingFailed("font glyph request size overflow".into()))?;
    let mut glyph_ids = Vec::new();
    glyph_ids
        .try_reserve_exact(capacity)
        .map_err(|source| FontError::Allocation {
            resource: "mapped font glyph IDs",
            source,
        })?;
    glyph_ids.push(0);

    for codepoint in codepoints.iter() {
        if char::from_u32(codepoint).is_none() {
            return Err(FontError::SubsettingFailed(format!(
                "font glyph request contains non-scalar Unicode value U+{codepoint:04X}"
            )));
        }
        if let Some(glyph_id) = subtable.map_glyph(codepoint).map_err(|error| {
            FontError::SubsettingFailed(format!("invalid cmap mapping: {error}"))
        })? && glyph_id != 0
        {
            glyph_ids.push(glyph_id);
        }
    }

    glyph_ids.sort_unstable();
    glyph_ids.dedup();
    Ok(glyph_ids)
}
