use std::marker::PhantomData;

use crate::big5::big5_to_unicode;
use crate::binary::budget;
use crate::binary::read::ReadScope;
use crate::error::ParseError;
use crate::font::Encoding;
use crate::macroman::{char_to_macroman, is_macroman, macroman_to_char};
use crate::subset::{CmapTarget, SubsetGlyphs};
use crate::tables::cmap::{Cmap, EncodingId, PlatformId, SequentialMapGroup, owned};
use crate::tables::os2::{self, Os2};
use crate::tables::{FontTableProvider, cmap};
use crate::tag;

pub struct MappingsToKeep<T> {
    // Keep mappings in a sorted vector instead of a tree. The vector's exact
    // source-derived capacity can be charged before allocation; tree nodes
    // would otherwise bypass the selected budget one allocation at a time.
    mappings: Vec<(Character, u16)>,
    plane: CharExistence,
    _ids: PhantomData<T>,
}

pub enum OldIds {}
pub enum NewIds {}

#[derive(Debug, Ord, PartialOrd, Eq, PartialEq, Copy, Clone)]
enum CharExistence {
    /// Can be encoded in [MacRoman](https://en.wikipedia.org/wiki/Mac_OS_Roman)
    MacRoman = 1,
    /// Unicode Plane 0
    BasicMultilingualPlane = 2,
    /// Unicode Plane 1 onwards
    AstralPlane = 3,
    /// Exists outside Unicode
    DivinePlane = 4,
}

#[derive(Debug, Copy, Clone, Ord, PartialOrd, Eq, PartialEq)]
enum Character {
    Unicode(char),
    Symbol(u32),
}

/// The strategy to use to generate a cmap table for the subset font
#[allow(unused)]
pub(crate) enum CmapStrategy {
    /// Build a cmap table by filtering existing mappings
    Generate(MappingsToKeep<OldIds>), // FIXME: Rename
    /// Use the supplied Mac Roman cmap table
    MacRomanSupplied(Box<[u8; 256]>),
    /// Omit the cmap table
    Omit,
}

#[derive(Debug)]
struct CmapSubtableFormat4Segment<'a> {
    start: u32,
    end: u32,
    glyph_ids: &'a mut Vec<u16>,
    consecutive_glyph_ids: bool,
}

impl Character {
    fn new(ch: u32, encoding: Encoding) -> Option<Self> {
        match encoding {
            Encoding::Unicode => std::char::from_u32(ch).map(Character::Unicode),
            Encoding::Symbol => Some(Character::Symbol(ch)),
            Encoding::AppleRoman => macroman_to_char(ch as u8).map(Character::Unicode),
            Encoding::Big5 => u16::try_from(ch)
                .ok()
                .and_then(big5_to_unicode)
                .map(Character::Unicode),
        }
    }

    fn existence(self) -> CharExistence {
        match self {
            Character::Unicode(ch) if is_macroman(ch) => CharExistence::MacRoman,
            Character::Unicode(ch) if ch <= '\u{FFFF}' => CharExistence::BasicMultilingualPlane,
            Character::Unicode(_) => CharExistence::AstralPlane,
            Character::Symbol(_) => CharExistence::DivinePlane,
        }
    }

    fn as_u32(self) -> u32 {
        match self {
            Character::Unicode(ch) => ch as u32,
            Character::Symbol(ch) => ch,
        }
    }
}

impl From<char> for Character {
    fn from(ch: char) -> Self {
        Character::Unicode(ch)
    }
}

impl<'a> CmapSubtableFormat4Segment<'a> {
    fn new(start: u32, gid: u16, glyph_ids: &'a mut Vec<u16>) -> Result<Self, ParseError> {
        glyph_ids.clear();
        budget::reserve_vec(glyph_ids, 1).map_err(|_| ParseError::LimitExceeded)?;
        glyph_ids.push(gid);
        Ok(CmapSubtableFormat4Segment {
            start,
            end: start,
            glyph_ids,
            consecutive_glyph_ids: true,
        })
    }

    fn add(&mut self, ch: u32, gid: u16) -> Result<bool, ParseError> {
        // -1 because the next consecutive character introduces no gap
        let gap = ch.saturating_sub(self.end).saturating_sub(1);
        let should_remain_compact = self.consecutive_glyph_ids && self.glyph_ids.len() >= 4;

        if gap > 0 && should_remain_compact {
            // Each new segment costs 8 bytes so if the gap will introduce a non-consecutive glyph
            // id and the current segment contains 4 of more entries it's better to start a new
            // segment and allow this one to continue to use the compact representation.
            Ok(false)
        } else if gap < 4 {
            // Each gap entry is two bytes in the glyph id array, if the gap is less than four
            // characters then it's worth adding to this segment, otherwise it's better to create
            // a new segment (which costs 8 bytes).

            // Gaps need to be mapped to .notdef (glyph id 0)
            if gap == 0 {
                // NOTE(unwrap): glyph_ids is never empty
                let prev = self.glyph_ids.last().copied().unwrap();
                self.consecutive_glyph_ids &= (prev + 1) == gid;
            } else {
                budget::reserve_vec(self.glyph_ids, gap as usize + 1)
                    .map_err(|_| ParseError::LimitExceeded)?;
                self.glyph_ids.extend(std::iter::repeat_n(0, gap as usize));
                // if there's a gap then the glyph ids can't be consecutive
                self.consecutive_glyph_ids = false;
            }
            budget::reserve_vec(self.glyph_ids, 1).map_err(|_| ParseError::LimitExceeded)?;
            self.glyph_ids.push(gid);
            self.end = ch;
            Ok(true)
        } else {
            Ok(false)
        }
    }
}

impl owned::CmapSubtableFormat4 {
    fn from_mappings(
        mappings: &MappingsToKeep<NewIds>,
    ) -> Result<owned::CmapSubtableFormat4, ParseError> {
        let mut table = owned::CmapSubtableFormat4 {
            language: 0,
            end_codes: Vec::new(),
            start_codes: Vec::new(),
            id_deltas: Vec::new(),
            id_range_offsets: Vec::new(),
            glyph_id_array: Vec::new(),
        };

        let segment_capacity = mappings.mappings.len().saturating_add(1);
        budget::reserve_vec(&mut table.end_codes, segment_capacity)
            .map_err(|_| ParseError::LimitExceeded)?;
        budget::reserve_vec(&mut table.start_codes, segment_capacity)
            .map_err(|_| ParseError::LimitExceeded)?;
        budget::reserve_vec(&mut table.id_deltas, segment_capacity)
            .map_err(|_| ParseError::LimitExceeded)?;
        budget::reserve_vec(&mut table.id_range_offsets, segment_capacity)
            .map_err(|_| ParseError::LimitExceeded)?;
        budget::reserve_vec(&mut table.glyph_id_array, mappings.mappings.len())
            .map_err(|_| ParseError::LimitExceeded)?;

        // Group the mappings into contiguous ranges, there can be holes in the ranges
        let mut glyph_ids = Vec::new();
        budget::reserve_vec(&mut glyph_ids, mappings.mappings.len())
            .map_err(|_| ParseError::LimitExceeded)?;
        let mut id_range_offset_fixup_indices = Vec::new();
        budget::reserve_vec(
            &mut id_range_offset_fixup_indices,
            mappings.mappings.len().saturating_add(1),
        )
        .map_err(|_| ParseError::LimitExceeded)?;
        if let Some((start, gid)) = mappings.iter().next() {
            let mut segment = CmapSubtableFormat4Segment::new(start.as_u32(), gid, &mut glyph_ids)?;
            for (ch, gid) in mappings.iter().skip(1) {
                if !segment.add(ch.as_u32(), gid)? {
                    table.add_segment(segment, &mut id_range_offset_fixup_indices)?;
                    segment = CmapSubtableFormat4Segment::new(ch.as_u32(), gid, &mut glyph_ids)?;
                }
            }

            // Add final range
            table.add_segment(segment, &mut id_range_offset_fixup_indices)?;
        }

        // Final start code and endCode values must be 0xFFFF. This segment need not contain any
        // valid mappings. (It can just map the single character code 0xFFFF to missingGlyph).
        // However, the segment must be present.
        //
        // — https://docs.microsoft.com/en-us/typography/opentype/spec/cmap#format-4-segment-mapping-to-delta-values
        let segment = CmapSubtableFormat4Segment::new(0xFFFF, 0, &mut glyph_ids)?;
        table.add_segment(segment, &mut id_range_offset_fixup_indices)?;

        // Fix up the id_range_offsets now that all segments have been added
        let num_segments = table.end_codes.len();
        for index in id_range_offset_fixup_indices {
            let id_range_offset = &mut table.id_range_offsets[index];
            let count = num_segments
                .checked_add(usize::from(*id_range_offset))
                .and_then(|count| count.checked_sub(index))
                .ok_or(ParseError::LimitExceeded)?;
            // ×2 because we need to skip over `count` 16-bit values
            let offset = count.checked_mul(2).ok_or(ParseError::LimitExceeded)?;
            *id_range_offset = u16::try_from(offset).map_err(|_| ParseError::LimitExceeded)?;
        }

        Ok(table)
    }

    fn add_segment(
        &mut self,
        segment: CmapSubtableFormat4Segment<'_>,
        id_range_offset_fixups: &mut Vec<usize>,
    ) -> Result<(), ParseError> {
        budget::reserve_vec(&mut self.start_codes, 1).map_err(|_| ParseError::LimitExceeded)?;
        budget::reserve_vec(&mut self.end_codes, 1).map_err(|_| ParseError::LimitExceeded)?;
        self.start_codes.push(segment.start as u16);
        self.end_codes.push(segment.end as u16);

        // If the segment contains contiguous range of glyph ids then we can just store
        // an id delta for the entire range.
        if segment.consecutive_glyph_ids {
            // NOTE(unwrap): safe as segments will always contain at least one char->glyph mapping
            let first_glyph_id = *segment.glyph_ids.first().unwrap();

            // NOTE: casting start to i32 is safe as format 4 can only hold Unicode BMP chars,
            // which are 16-bit values. Casting the result to i16 is safe because the calculation
            // is modulo 0x10000 (65536), which limits the value to ±0x10000.
            budget::reserve_vec(&mut self.id_deltas, 1).map_err(|_| ParseError::LimitExceeded)?;
            self.id_deltas
                .push((i32::from(first_glyph_id) - segment.start as i32 % 0x10000) as i16);
            budget::reserve_vec(&mut self.id_range_offsets, 1)
                .map_err(|_| ParseError::LimitExceeded)?;
            self.id_range_offsets.push(0);
        } else {
            budget::reserve_vec(&mut self.id_deltas, 1).map_err(|_| ParseError::LimitExceeded)?;
            // Glyph ids are not consecutive so store them in the glyph id array via id range
            // offsets
            self.id_deltas.push(0);
            // NOTE: The id range offset value will be fixed up in a later pass
            budget::reserve_vec(id_range_offset_fixups, 1)
                .map_err(|_| ParseError::LimitExceeded)?;
            id_range_offset_fixups.push(self.id_range_offsets.len());
            // NOTE: casting should be safe as num_glyphs in a font is u16
            budget::reserve_vec(&mut self.id_range_offsets, 1)
                .map_err(|_| ParseError::LimitExceeded)?;
            self.id_range_offsets.push(self.glyph_id_array.len() as u16);
            budget::reserve_vec(&mut self.glyph_id_array, segment.glyph_ids.len())
                .map_err(|_| ParseError::LimitExceeded)?;
            self.glyph_id_array.extend_from_slice(segment.glyph_ids);
        }
        Ok(())
    }
}

impl owned::CmapSubtableFormat12 {
    fn from_mappings(
        mappings: &MappingsToKeep<NewIds>,
    ) -> Result<owned::CmapSubtableFormat12, ParseError> {
        let mut segments = Vec::new();
        budget::reserve_vec(&mut segments, mappings.mappings.len())
            .map_err(|_| ParseError::LimitExceeded)?;
        if let Some((start, gid)) = mappings.iter().next() {
            let mut segment = SequentialMapGroup {
                start_char_code: start.as_u32(),
                end_char_code: start.as_u32(),
                start_glyph_id: u32::from(gid),
            };
            let mut prev_gid = gid;
            for (ch, gid) in mappings.iter().skip(1) {
                if ch.as_u32() == segment.end_char_code + 1 && gid == prev_gid + 1 {
                    segment.end_char_code += 1
                } else {
                    segments.push(segment);
                    segment = SequentialMapGroup {
                        start_char_code: ch.as_u32(),
                        end_char_code: ch.as_u32(),
                        start_glyph_id: u32::from(gid),
                    };
                }
                prev_gid = gid;
            }
            segments.push(segment);
        }

        Ok(owned::CmapSubtableFormat12 {
            language: 0,
            groups: segments,
        })
    }
}

impl owned::EncodingRecord {
    pub fn from_mappings(mappings: &MappingsToKeep<NewIds>) -> Result<Self, ParseError> {
        match mappings.plane() {
            CharExistence::MacRoman => {
                // The language field must be set to zero for all 'cmap' subtables whose platform
                // IDs are other than Macintosh (platform ID 1). For 'cmap' subtables whose
                // platform IDs are Macintosh, set this field to the Macintosh language ID of the
                // 'cmap' subtable plus one, or to zero if the 'cmap' subtable is not
                // language-specific. For example, a Mac OS Turkish 'cmap' subtable must set this
                // field to 18, since the Macintosh language ID for Turkish is 17. A Mac OS Roman
                // 'cmap' subtable must set this field to 0, since Mac OS Roman is not a
                // language-specific encoding.
                //
                // — https://docs.microsoft.com/en-us/typography/opentype/spec/cmap#use-of-the-language-field-in-cmap-subtables
                let mut glyph_id_array = [0; 256];
                for (ch, gid) in mappings.iter() {
                    let ch_mac = match ch {
                        // NOTE(unwrap): Safe as we verified all chars with `is_macroman` earlier
                        Character::Unicode(unicode) => {
                            usize::from(char_to_macroman(unicode).unwrap())
                        },
                        Character::Symbol(_) => unreachable!("symbol in mac roman"),
                    };
                    // Cast is safe as we determined that all chars are valid in Mac Roman
                    glyph_id_array[ch_mac] = gid as u8;
                }
                budget::charge_allocation(std::mem::size_of::<[u8; 256]>())
                    .map_err(|_| ParseError::LimitExceeded)?;
                let sub_table = owned::CmapSubtable::Format0 {
                    language: 0,
                    glyph_id_array: Box::new(glyph_id_array),
                };
                Ok(owned::EncodingRecord {
                    platform_id: PlatformId::MACINTOSH,
                    encoding_id: EncodingId::MACINTOSH_APPLE_ROMAN,
                    sub_table,
                })
            },
            CharExistence::BasicMultilingualPlane => {
                let sub_table = cmap::owned::CmapSubtable::Format4(
                    owned::CmapSubtableFormat4::from_mappings(mappings)?,
                );
                Ok(owned::EncodingRecord {
                    platform_id: PlatformId::UNICODE,
                    encoding_id: EncodingId::UNICODE_BMP,
                    sub_table,
                })
            },
            CharExistence::AstralPlane => {
                let sub_table = cmap::owned::CmapSubtable::Format12(
                    owned::CmapSubtableFormat12::from_mappings(mappings)?,
                );
                Ok(owned::EncodingRecord {
                    platform_id: PlatformId::UNICODE,
                    encoding_id: EncodingId::UNICODE_FULL,
                    sub_table,
                })
            },
            CharExistence::DivinePlane => {
                let sub_table = cmap::owned::CmapSubtable::Format4(
                    owned::CmapSubtableFormat4::from_mappings(mappings)?,
                );
                Ok(owned::EncodingRecord {
                    platform_id: PlatformId::WINDOWS,
                    encoding_id: EncodingId::WINDOWS_SYMBOL,
                    sub_table,
                })
            },
        }
    }
}

impl<T> MappingsToKeep<T> {
    fn iter(&self) -> impl Iterator<Item = (Character, u16)> + '_ {
        self.mappings.iter().copied()
    }

    fn plane(&self) -> CharExistence {
        self.plane
    }
}

impl MappingsToKeep<OldIds> {
    pub(crate) fn new(
        provider: &impl FontTableProvider,
        glyph_ids: &[u16],
        target: CmapTarget,
    ) -> Result<Self, ParseError> {
        let cmap_data = provider.read_table_data(tag::CMAP)?;
        let cmap0 = ReadScope::new(&cmap_data).read::<Cmap<'_>>()?;
        let (encoding, cmap_sub_table) =
            crate::font::read_cmap_subtable(&cmap0)?.ok_or(ParseError::UnsuitableCmap)?;

        // Special case handling of a Symbol cmap targeting MacRoman
        //
        // This exists to handle the case where MacRoman characters were successfully mapped to
        // glyphs via a cmap sub-table with Symbol encoding (see `legacy_symbol_char_code` in
        // `Font`). We need to perform the inverse operation when we're targeting a MacRoman
        // encoded cmap sub-table in the subset font.
        let symbol_first_char = if encoding == Encoding::Symbol && target == CmapTarget::MacRoman {
            Some(
                provider
                    .table_data(tag::OS_2)?
                    .map(|data| ReadScope::new(&data).read_dep::<Os2>(data.len()))
                    .transpose()?
                    .map(|os2| os2.us_first_char_index)
                    .unwrap_or(0x20),
            )
        } else {
            None
        };

        // Collect cmap mappings for the selected glyph ids
        // Keep the source order beside each mapping. The old BTreeMap-based
        // implementation replaced duplicate characters with the last source
        // mapping; an unstable sort must carry that order explicitly to keep
        // the same fidelity.
        let mut mappings_to_keep: Vec<(Character, u16, usize)> = Vec::new();
        budget::reserve_vec(&mut mappings_to_keep, glyph_ids.len())
            .map_err(|_| ParseError::LimitExceeded)?;
        let mut allocation_error = None;
        let mut source_order = 0usize;
        let mut plane = if target == CmapTarget::Unicode {
            CharExistence::BasicMultilingualPlane
        } else {
            CharExistence::MacRoman
        };

        // Process all the mappings and select the ones we want to keep
        cmap_sub_table.mappings_fn(|ch, gid| {
            if allocation_error.is_some() {
                return;
            }
            if gid != 0 && glyph_ids.contains(&gid) {
                // We want to keep this mapping, determine the plane it lives on
                // If `symbol_first_char` is set then that indicates we're targeting a MacRoman
                // cmap sub-table with a source Windows Symbol encoded cmap sub-table. Perform
                // a mapping from symbol char code to unicode (inverse of what was done when
                // mapping glyphs).
                let output_char = symbol_first_char
                    .and_then(|first| legacy_symbol_char_code_to_unicode(ch, first))
                    .map(|uni| Some(Character::from(uni)))
                    .unwrap_or_else(|| Character::new(ch, encoding));
                let output_char = match output_char {
                    Some(ch) => ch,
                    None => return,
                };

                match target {
                    CmapTarget::MacRoman => {
                        // Only keep if it's MacRoman compatible
                        if output_char.existence() <= CharExistence::MacRoman {
                            if budget::reserve_for_push(&mut mappings_to_keep).is_err() {
                                allocation_error = Some(ParseError::LimitExceeded);
                                return;
                            }
                            let order = match source_order.checked_add(1) {
                                Some(next) => {
                                    let order = source_order;
                                    source_order = next;
                                    order
                                },
                                None => {
                                    allocation_error = Some(ParseError::LimitExceeded);
                                    return;
                                },
                            };
                            mappings_to_keep.push((output_char, gid, order));
                        }
                    },
                    CmapTarget::Unicode | CmapTarget::Unrestricted => {
                        if output_char.existence() > plane {
                            plane = output_char.existence();
                        }
                        if budget::reserve_for_push(&mut mappings_to_keep).is_err() {
                            allocation_error = Some(ParseError::LimitExceeded);
                            return;
                        }
                        let order = match source_order.checked_add(1) {
                            Some(next) => {
                                let order = source_order;
                                source_order = next;
                                order
                            },
                            None => {
                                allocation_error = Some(ParseError::LimitExceeded);
                                return;
                            },
                        };
                        mappings_to_keep.push((output_char, gid, order));
                    },
                }
            }
        })?;

        if let Some(error) = allocation_error {
            return Err(error);
        }

        let mappings = retain_last_mappings(mappings_to_keep)?;

        if mappings.len() <= usize::from(u16::MAX) {
            Ok(MappingsToKeep {
                mappings,
                plane,
                _ids: PhantomData,
            })
        } else {
            Err(ParseError::LimitExceeded)
        }
    }

    /// Update the glyph ids to be ids in the new subset font
    pub(crate) fn update_to_new_ids(
        mut self,
        subset_glyphs: &impl SubsetGlyphs,
    ) -> MappingsToKeep<NewIds> {
        self.mappings
            .iter_mut()
            .for_each(|(_ch, gid)| *gid = subset_glyphs.new_id(*gid));
        MappingsToKeep {
            mappings: self.mappings,
            plane: self.plane,
            _ids: PhantomData,
        }
    }

    // Calculate new first and last Unicode codepoints
    pub(crate) fn first_last_codepoints(&self) -> (u32, u32) {
        if self.mappings.is_empty() {
            (0, 0) // No mappings, use 0 for both
        } else {
            self.iter().fold((u32::MAX, 0_u32), |(min, max), (ch, _)| {
                let code = ch.as_u32();
                (min.min(code), max.max(code))
            })
        }
    }

    // Compute the new OS/2 ulUnicodeRange bitmask
    pub(crate) fn unicode_bitmask(&self) -> u128 {
        self.iter().fold(0, |mask, (ch, _)| {
            mask | os2::unicode_range_mask(ch.as_u32())
        })
    }
}

// See also legacy_symbol_char_code in Font and the explanation there
fn legacy_symbol_char_code_to_unicode(ch: u32, first_char: u16) -> Option<char> {
    let char_code0 = if (0xF000..=0xF0FF).contains(&ch) {
        ch
    } else {
        ch + 0xF000
    };
    std::char::from_u32((char_code0 + 0x20) - u32::from(first_char)) // Perform subtraction last to avoid underflow.
}

fn retain_last_mappings(
    mut mappings_to_keep: Vec<(Character, u16, usize)>,
) -> Result<Vec<(Character, u16)>, ParseError> {
    mappings_to_keep.sort_unstable_by_key(|(character, _, order)| (*character, *order));
    let mut unique = 0;
    for index in 0..mappings_to_keep.len() {
        if index == 0 || mappings_to_keep[index].0 != mappings_to_keep[unique - 1].0 {
            mappings_to_keep.swap(unique, index);
            unique += 1;
        } else {
            // BTreeMap::insert replaced a duplicate key with the latest
            // mapping; preserve that behaviour after sorting by source order.
            mappings_to_keep[unique - 1].1 = mappings_to_keep[index].1;
        }
    }
    mappings_to_keep.truncate(unique);

    let mut mappings = Vec::new();
    budget::reserve_vec(&mut mappings, mappings_to_keep.len())
        .map_err(|_| ParseError::LimitExceeded)?;
    mappings.extend(
        mappings_to_keep
            .into_iter()
            .map(|(character, gid, _order)| (character, gid)),
    );
    Ok(mappings)
}

#[cfg(test)]
mod tests {
    use crate::tables::OpenTypeFont;
    use crate::tests::read_fixture;

    use super::*;

    #[test]
    fn test_character_existence() {
        assert_eq!(Character::Unicode('a').existence(), CharExistence::MacRoman);
        assert_eq!(
            Character::Unicode('ռ').existence(),
            CharExistence::BasicMultilingualPlane
        );
        assert_eq!(
            Character::Unicode('🦀').existence(),
            CharExistence::AstralPlane
        );
    }

    #[test]
    fn duplicate_character_keeps_last_source_mapping() {
        let mappings = retain_last_mappings(vec![
            (Character::Unicode('a'), 3, 0),
            (Character::Unicode('b'), 4, 1),
            (Character::Unicode('a'), 9, 2),
        ])
        .unwrap();

        assert_eq!(
            mappings,
            vec![(Character::Unicode('a'), 9), (Character::Unicode('b'), 4)]
        );
    }

    #[test]
    fn test_format4_subtable() {
        let mappings = MappingsToKeep {
            mappings: vec![
                (Character::Unicode('a'), 1),
                (Character::Unicode('b'), 2),
                (Character::Unicode('i'), 4),
                (Character::Unicode('j'), 3),
            ]
            .into_iter()
            .collect(),
            plane: CharExistence::MacRoman,
            _ids: PhantomData,
        };
        let sub_table = owned::CmapSubtableFormat4::from_mappings(&mappings).unwrap();
        let expected = owned::CmapSubtableFormat4 {
            language: 0,
            start_codes: vec![97, 105, 0xFFFF],
            end_codes: vec![98, 106, 0xFFFF],
            id_deltas: vec![-96, 0, 1],
            id_range_offsets: vec![0, 4, 0],
            glyph_id_array: vec![4, 3],
        };
        assert_eq!(sub_table, expected);
    }

    #[test]
    fn test_format12_subtable() {
        let mappings = MappingsToKeep {
            mappings: vec![
                (Character::Unicode('a'), 1),
                (Character::Unicode('b'), 2),
                (Character::Unicode('🦀'), 3),
                (Character::Unicode('🦁'), 4),
            ]
            .into_iter()
            .collect(),
            plane: CharExistence::AstralPlane,
            _ids: PhantomData,
        };
        let sub_table = owned::CmapSubtableFormat12::from_mappings(&mappings).unwrap();
        let expected = owned::CmapSubtableFormat12 {
            language: 0,
            groups: vec![
                SequentialMapGroup {
                    start_char_code: 97,
                    end_char_code: 98,
                    start_glyph_id: 1,
                },
                SequentialMapGroup {
                    start_char_code: 129408,
                    end_char_code: 129409,
                    start_glyph_id: 3,
                },
            ],
        };
        assert_eq!(sub_table, expected);
    }

    #[test]
    fn test_format12_subtable_empty_mappings() {
        let mappings = MappingsToKeep {
            mappings: Vec::new(),
            plane: CharExistence::BasicMultilingualPlane,
            _ids: PhantomData,
        };
        let sub_table = owned::CmapSubtableFormat12::from_mappings(&mappings).unwrap();
        assert!(sub_table.groups.is_empty());
    }

    #[test]
    fn test_target_macroman_from_symbol() {
        let buffer = read_fixture("tests/fonts/opentype/SymbolTest-Regular.ttf");
        let scope = ReadScope::new(&buffer);
        let font_file = scope
            .read::<OpenTypeFont<'_>>()
            .expect("unable to parse font file");
        let table_provider = font_file
            .table_provider(0)
            .expect("unable to create font provider");

        let to_keep = MappingsToKeep::new(&table_provider, &[0, 3, 4, 5], CmapTarget::MacRoman)
            .expect("error building mappings to keep");
        assert_eq!(to_keep.plane, CharExistence::MacRoman);
        let chars: String = to_keep
            .mappings
            .iter()
            .map(|(ch, _gid)| match ch {
                Character::Unicode(c) => *c,
                Character::Symbol(_) => panic!("expected Character::Unicode got Character::Symbol"),
            })
            .collect();
        assert_eq!(chars, "abc");
    }
}
