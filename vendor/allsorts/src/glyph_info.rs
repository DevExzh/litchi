#![deny(missing_docs)]

//! Utilities for accessing glyph information such as advance.

use std::borrow::Cow;
use std::fmt::Write as _;
use std::mem::size_of;
use std::sync::Arc;

use ouroboros::self_referencing;

use crate::GlyphId;
use crate::binary::budget;
use crate::binary::read::ReadScope;
use crate::error::ParseError;
use crate::font::Encoding;
use crate::macroman::macroman_to_char;
use crate::post::PostTable;
use crate::tables::cmap::CmapSubtable;
use crate::tables::{HheaTable, HmtxTable, MaxpTable};

/// Retrieve glyph advance.
///
/// Since the `hhea` and `vhea` tables share the same format this function will return horizontal
/// or vertical advance depending on whether `hhea` or `vhea` is supplied to the `hhea` argument.
pub fn advance(
    maxp: &MaxpTable,
    hhea: &HheaTable,
    hmtx_data: &[u8],
    glyph: GlyphId,
) -> Result<u16, ParseError> {
    // Avoid parsing hmtx in this case
    if i32::from(glyph) > i32::from(maxp.num_glyphs) - 1 {
        return Ok(0);
    }

    let glyph = usize::from(glyph);
    let num_glyphs = usize::from(maxp.num_glyphs);
    let num_metrics = usize::from(hhea.num_h_metrics);
    let hmtx = ReadScope::new(hmtx_data).read_dep::<HmtxTable<'_>>((num_glyphs, num_metrics))?;

    if glyph < num_metrics {
        Ok(hmtx
            .h_metrics
            .get_item(glyph)
            .map_or(0, |x| x.advance_width))
    } else if num_metrics > 0 {
        Ok(hmtx
            .h_metrics
            .get_item(num_metrics - 1)
            .map_or(0, |metrics| metrics.advance_width))
    } else {
        Err(ParseError::BadIndex)
    }
}

#[self_referencing]
struct Post {
    data: Box<[u8]>,
    #[borrows(data)]
    #[not_covariant]
    table: PostTable<'this>,
}

/// Structure for looking up glyph names.
pub struct GlyphNames {
    post: Option<Post>,
    cmap: Option<CmapMappings>,
}

struct CmapMappings {
    encoding: Encoding,
    mappings: Vec<CmapMapping>,
}

#[derive(Clone, Copy)]
struct CmapMapping {
    glyph: u16,
    character: u32,
    order: usize,
}

impl GlyphNames {
    /// Construct a new `GlyphNames` instance.
    pub fn new(
        cmap_subtable: &Option<(Encoding, CmapSubtable<'_>)>,
        post_data: Option<Box<[u8]>>,
    ) -> Self {
        let post = post_data.and_then(|data| {
            PostTryBuilder {
                data,
                table_builder: |data| ReadScope::new(data).read::<PostTable<'_>>(),
            }
            .try_build()
            .ok()
        });
        let cmap = cmap_subtable
            .as_ref()
            .and_then(|(encoding, subtable)| CmapMappings::new(*encoding, subtable).ok().flatten());
        GlyphNames { post, cmap }
    }

    /// Construct glyph-name data while propagating a selected allocation
    /// budget. The infallible constructor above remains for existing callers.
    pub fn try_new(
        cmap_subtable: &Option<(Encoding, CmapSubtable<'_>)>,
        post_data: Option<Box<[u8]>>,
    ) -> Result<Self, ParseError> {
        let post = post_data.and_then(|data| {
            PostTryBuilder {
                data,
                table_builder: |data| ReadScope::new(data).read::<PostTable<'_>>(),
            }
            .try_build()
            .ok()
        });
        if budget::failure().is_some() {
            return Err(ParseError::LimitExceeded);
        }
        if budget::failed() {
            return Err(ParseError::BadValue);
        }
        let cmap = cmap_subtable
            .as_ref()
            .map(|(encoding, subtable)| CmapMappings::new(*encoding, subtable))
            .transpose()?
            .flatten();
        Ok(GlyphNames { post, cmap })
    }

    /// Look up the name of `gid` in the `post` and `cmap` tables.
    pub fn glyph_name<'a>(&self, gid: GlyphId) -> Cow<'a, str> {
        // Glyph 0 is always .notdef
        if gid == 0 {
            return Cow::from(".notdef");
        }

        self.glyph_name_from_post(gid)
            .or_else(|| self.glyph_name_from_cmap(gid))
            .unwrap_or_else(|| Cow::from(format!("g{}", gid)))
    }

    fn try_glyph_name<'a>(&self, gid: GlyphId) -> Result<Cow<'a, str>, ParseError> {
        if gid == 0 {
            return Ok(Cow::from(".notdef"));
        }

        if let Some(name) = self.try_glyph_name_from_post(gid)? {
            return Ok(name);
        }
        if let Some(name) = self.glyph_name_from_cmap(gid) {
            return Ok(name);
        }

        let mut fallback = String::new();
        let digits = decimal_digits(usize::from(gid));
        let required = 1usize
            .checked_add(digits)
            .ok_or(ParseError::LimitExceeded)?;
        budget::reserve_string(&mut fallback, required).map_err(|_| ParseError::LimitExceeded)?;
        write!(&mut fallback, "g{gid}").map_err(|_| ParseError::BadValue)?;
        Ok(Cow::Owned(fallback))
    }

    /// Determine the set of unique glyph names for the supplied glyph ids.
    pub fn unique_glyph_names<'a>(&self, ids: &[GlyphId]) -> Vec<Cow<'a, str>> {
        self.try_unique_glyph_names(ids).unwrap_or_default()
    }

    /// Determine unique glyph names while propagating allocation failures.
    pub fn try_unique_glyph_names<'a>(
        &self,
        ids: &[GlyphId],
    ) -> Result<Vec<Cow<'a, str>>, ParseError> {
        let mut names = Vec::new();
        budget::reserve_vec(&mut names, ids.len()).map_err(|_| ParseError::LimitExceeded)?;
        for &gid in ids {
            names.push(self.try_glyph_name(gid)?);
        }
        unique_glyph_names_budgeted(names.into_iter(), ids.len())
    }

    fn glyph_name_from_post<'a>(&self, gid: GlyphId) -> Option<Cow<'a, str>> {
        let post = self.post.as_ref()?;
        post.glyph_name(gid)
    }

    fn try_glyph_name_from_post<'a>(
        &self,
        gid: GlyphId,
    ) -> Result<Option<Cow<'a, str>>, ParseError> {
        let Some(post) = self.post.as_ref() else {
            return Ok(None);
        };
        post.try_glyph_name(gid)
    }

    fn glyph_name_from_cmap<'a>(&self, gid: GlyphId) -> Option<Cow<'a, str>> {
        let cmap = self.cmap.as_ref()?;
        cmap.glyph_name(gid)
    }
}

fn unique_glyph_names_budgeted<'a>(
    names: impl Iterator<Item = Cow<'a, str>>,
    capacity: usize,
) -> Result<Vec<Cow<'a, str>>, ParseError> {
    let mut source: Vec<Arc<Cow<'a, str>>> = Vec::new();
    budget::reserve_vec(&mut source, capacity).map_err(|_| ParseError::LimitExceeded)?;
    for name in names {
        budget::reserve_for_push(&mut source).map_err(|_| ParseError::LimitExceeded)?;
        let arc_bytes = size_of::<Cow<'a, str>>()
            .checked_add(2 * size_of::<usize>())
            .ok_or(ParseError::LimitExceeded)?;
        budget::charge_allocation(arc_bytes).map_err(|_| ParseError::LimitExceeded)?;
        source.push(Arc::new(name));
    }

    // Sort indexes instead of inserting into a sorted map. This keeps the
    // budgeted path O(n log n) for large glyph sets while retaining the
    // original input order for the `.altNN` suffix.
    let mut order = Vec::new();
    budget::reserve_vec(&mut order, source.len()).map_err(|_| ParseError::LimitExceeded)?;
    for index in 0..source.len() {
        budget::reserve_for_push(&mut order).map_err(|_| ParseError::LimitExceeded)?;
        order.push(index);
    }
    order.sort_unstable_by(|left, right| {
        let left_name: &str = source[*left].as_ref();
        let right_name: &str = source[*right].as_ref();
        left_name.cmp(right_name).then(left.cmp(right))
    });

    let mut unique_names: Vec<(usize, Arc<Cow<'a, str>>)> = Vec::new();
    budget::reserve_vec(&mut unique_names, source.len()).map_err(|_| ParseError::LimitExceeded)?;

    let mut group_start = 0;
    while group_start < order.len() {
        let first_index = order[group_start];
        let first_name: &str = source[first_index].as_ref();
        let mut group_end = group_start + 1;
        while group_end < order.len() {
            let name: &str = source[order[group_end]].as_ref();
            if name != first_name {
                break;
            }
            group_end += 1;
        }
        for (alt, &index) in order[group_start..group_end].iter().enumerate() {
            let name = Arc::clone(&source[index]);
            let unique_name = if alt == 0 {
                name
            } else {
                let mut generated = String::new();
                let suffix_digits = decimal_digits(alt).max(2);
                let required = first_name
                    .len()
                    .checked_add(4)
                    .and_then(|length| length.checked_add(suffix_digits))
                    .ok_or(ParseError::LimitExceeded)?;
                budget::reserve_string(&mut generated, required)
                    .map_err(|_| ParseError::LimitExceeded)?;
                write!(&mut generated, "{}.alt{:02}", first_name, alt)
                    .map_err(|_| ParseError::BadValue)?;
                let arc_bytes = size_of::<Cow<'a, str>>()
                    .checked_add(2 * size_of::<usize>())
                    .ok_or(ParseError::LimitExceeded)?;
                budget::charge_allocation(arc_bytes).map_err(|_| ParseError::LimitExceeded)?;
                Arc::new(Cow::Owned(generated))
            };
            budget::reserve_for_push(&mut unique_names).map_err(|_| ParseError::LimitExceeded)?;
            unique_names.push((index, unique_name));
        }
        group_start = group_end;
    }
    unique_names.sort_unstable_by_key(|(index, _)| *index);
    drop(order);
    drop(source);

    let mut result = Vec::new();
    budget::reserve_vec(&mut result, unique_names.len()).map_err(|_| ParseError::LimitExceeded)?;
    for (_, name) in unique_names {
        result.push(Arc::try_unwrap(name).map_err(|_| ParseError::BadValue)?);
    }
    Ok(result)
}

fn decimal_digits(mut value: usize) -> usize {
    let mut digits = 1;
    while value >= 10 {
        value /= 10;
        digits += 1;
    }
    digits
}

impl Post {
    fn glyph_name<'a>(&self, gid: GlyphId) -> Option<Cow<'a, str>> {
        self.with_table(|post: &PostTable<'_>| {
            match post.glyph_name(gid) {
                Ok(Some(glyph_name)) if glyph_name != ".notdef" => {
                    // Doesn't seem possible to avoid this allocation
                    Some(Cow::from(glyph_name.to_owned()))
                },
                _ => None,
            }
        })
    }

    fn try_glyph_name<'a>(&self, gid: GlyphId) -> Result<Option<Cow<'a, str>>, ParseError> {
        self.with_table(|post: &PostTable<'_>| match post.glyph_name(gid) {
            Ok(Some(glyph_name)) if glyph_name != ".notdef" => {
                let mut owned = String::new();
                budget::reserve_string(&mut owned, glyph_name.len())
                    .map_err(|_| ParseError::LimitExceeded)?;
                owned.push_str(glyph_name);
                Ok(Some(Cow::Owned(owned)))
            },
            Ok(_) => Ok(None),
            Err(ParseError::LimitExceeded) => Err(ParseError::LimitExceeded),
            Err(_) => Ok(None),
        })
    }
}

impl CmapMappings {
    fn new(
        encoding: Encoding,
        subtable: &CmapSubtable<'_>,
    ) -> Result<Option<CmapMappings>, ParseError> {
        let mut mappings = Vec::new();
        budget::reserve_vec(&mut mappings, subtable.size_hint())
            .map_err(|_| ParseError::LimitExceeded)?;
        let mut order = 0usize;
        let mut allocation_error = None;
        subtable.mappings_fn(|character, glyph| {
            if allocation_error.is_some() {
                return;
            }
            if budget::reserve_for_push(&mut mappings).is_err() {
                allocation_error = Some(ParseError::LimitExceeded);
                return;
            }
            let Some(next_order) = order.checked_add(1) else {
                allocation_error = Some(ParseError::LimitExceeded);
                return;
            };
            mappings.push(CmapMapping {
                glyph,
                character,
                order,
            });
            order = next_order;
        })?;
        if let Some(error) = allocation_error {
            return Err(error);
        }
        mappings.sort_unstable_by_key(|mapping| (mapping.glyph, mapping.order));
        let mut unique = 0;
        for index in 0..mappings.len() {
            if index == 0 || mappings[index].glyph != mappings[index - 1].glyph {
                mappings.swap(unique, index);
                unique += 1;
            }
        }
        mappings.truncate(unique);
        Ok(Some(CmapMappings { encoding, mappings }))
    }

    fn glyph_name<'a>(&self, gid: GlyphId) -> Option<Cow<'a, str>> {
        let index = self
            .mappings
            .binary_search_by_key(&gid, |mapping| mapping.glyph)
            .ok()?;
        let ch = self.mappings[index].character;
        match self.encoding {
            Encoding::AppleRoman => glyph_names::glyph_name(macroman_to_unicode(ch)?),
            Encoding::Unicode => glyph_names::glyph_name(ch),
            Encoding::Symbol => None,
            Encoding::Big5 => None, // FIXME
        }
    }
}

fn macroman_to_unicode(ch: u32) -> Option<u32> {
    macroman_to_char(ch as u8).map(|ch| ch as u32)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_unique_glyph_names() {
        let names = vec!["A"; 3].into_iter().map(Cow::from);
        let unique_names = unique_glyph_names_budgeted(names, 3).unwrap();
        assert_eq!(
            unique_names,
            &[Cow::from("A"), Cow::from("A.alt01"), Cow::from("A.alt02")]
        );
    }
}
