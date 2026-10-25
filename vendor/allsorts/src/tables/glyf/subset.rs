use super::{GlyfRecord, GlyfTable, Glyph, ParseError};
use crate::binary::budget;
use crate::subset::SubsetGlyphs;
use crate::tables::glyf::CompositeGlyph;

/// A subset glyph
#[derive(Clone)]
pub struct SubsetGlyph<'a> {
    /// The old glyph id of this glyph
    pub old_id: u16,
    /// The glyph
    pub record: GlyfRecord<'a>,
}

/// A `glyf` table that has been subset.
#[derive(Clone)]
pub struct SubsetGlyf<'a> {
    glyphs: Vec<SubsetGlyph<'a>>,
    /// Maps an old glyph index to its index in the new table
    old_to_new_id: Vec<u16>,
}

impl<'a> GlyfTable<'a> {
    /// Returns a copy of this table that only contains the glyphs specified by `glyph_ids`.
    pub fn subset(&self, glyph_ids: &[u16]) -> Result<SubsetGlyf<'a>, ParseError> {
        let mut selected_glyph_ids = Vec::new();
        budget::reserve_vec(&mut selected_glyph_ids, glyph_ids.len())
            .map_err(|_| ParseError::LimitExceeded)?;
        selected_glyph_ids.extend_from_slice(glyph_ids);
        let mut records = Vec::new();
        budget::reserve_vec(&mut records, glyph_ids.len())
            .map_err(|_| ParseError::LimitExceeded)?;

        let mut i = 0;
        while i < selected_glyph_ids.len() {
            let glyph_id = selected_glyph_ids[i];
            let source_record = self
                .records
                .get(usize::from(glyph_id))
                .ok_or(ParseError::BadIndex)?;
            super::charge_record_clone(source_record)?;
            let mut record = source_record.clone();
            if record.is_composite() {
                record.parse()?;
                let GlyfRecord::Parsed(Glyph::Composite(composite)) = &mut record else {
                    unreachable!("not a composite glyph")
                };
                add_glyph(&mut selected_glyph_ids, composite)?;
            }
            budget::reserve_vec(&mut records, 1).map_err(|_| ParseError::LimitExceeded)?;
            records.push(SubsetGlyph {
                old_id: glyph_id,
                record,
            });
            i += 1;
        }
        // Cast should be safe as there must be less than u16::MAX glyphs in a font
        let mut old_to_new_id = Vec::new();
        budget::reserve_vec(&mut old_to_new_id, self.records.len())
            .map_err(|_| ParseError::LimitExceeded)?;
        old_to_new_id.resize(self.records.len(), 0);
        for (new_id, glyph) in records.iter().enumerate() {
            old_to_new_id[usize::from(glyph.old_id)] = new_id as u16;
        }
        Ok(SubsetGlyf {
            glyphs: records,
            old_to_new_id,
        })
    }
}

impl SubsetGlyphs for SubsetGlyf<'_> {
    fn len(&self) -> usize {
        self.glyphs.len()
    }

    fn old_id(&self, new_id: u16) -> u16 {
        self.glyphs[usize::from(new_id)].old_id
    }

    fn new_id(&self, old_id: u16) -> u16 {
        self.old_to_new_id
            .get(usize::from(old_id))
            .copied()
            .unwrap_or(0)
    }
}

impl<'a> SubsetGlyf<'a> {
    pub(crate) fn into_glyf_table(self) -> Result<GlyfTable<'a>, ParseError> {
        let mut records = Vec::new();
        budget::reserve_vec(&mut records, self.glyphs.len())
            .map_err(|_| ParseError::LimitExceeded)?;
        records.extend(
            self.glyphs
                .into_iter()
                .map(|subset_record| subset_record.record),
        );
        Ok(GlyfTable { records })
    }
}

/// Add each of the child glyphs contained within a composite glyph to the subset font.
///
/// Updates the composite glyph indexes to point at the new child indexes.
fn add_glyph(glyph_ids: &mut Vec<u16>, composite: &mut CompositeGlyph) -> Result<(), ParseError> {
    for composite_glyph in composite.glyphs.iter_mut() {
        let new_id = match glyph_ids
            .iter()
            .position(|&id| id == composite_glyph.glyph_index)
        {
            Some(new_id) => new_id,
            None => {
                // Add this glyph to the list of ids to include in the subset font
                let new_id = glyph_ids.len();
                budget::reserve_vec(glyph_ids, 1).map_err(|_| ParseError::LimitExceeded)?;
                glyph_ids.push(composite_glyph.glyph_index);
                new_id
            },
        };
        composite_glyph.glyph_index = new_id as u16;
    }
    Ok(())
}
