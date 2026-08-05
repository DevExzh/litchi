//! Format-independent font preparation and publication helpers.

pub mod powerpoint;

use crate::{
    discovery::Loader,
    model::{FontData, FontError, FontProperties, GlyphMap, Style},
    subset::{Allsorts, Subsetter, glyph_ids},
};

/// One discovered and optionally subsetted, but still un-obfuscated, program.
#[derive(Debug)]
pub struct Prepared {
    pub name: String,
    pub style: Style,
    pub data: Vec<u8>,
    pub properties: FontProperties,
    pub subsetted: bool,
}

/// Discover and optionally subset fonts without knowing any package format.
///
/// The returned buffers are owned and un-obfuscated. Publication, content
/// types, relationship IDs, and part names belong to the document adapters.
pub fn prepare(used_glyphs: GlyphMap, subset_requested: bool) -> Result<Vec<Prepared>, FontError> {
    let loader = Loader::new();
    let subsetter = Allsorts::new();
    let mut requests = Vec::new();
    requests
        .try_reserve_exact(used_glyphs.len())
        .map_err(|source| FontError::Allocation {
            resource: "embedded-font requests",
            source,
        })?;
    requests.extend(used_glyphs);
    requests.sort_by(|(left, _), (right, _)| left.cmp(right));

    let mut prepared = Vec::new();
    prepared
        .try_reserve(requests.len())
        .map_err(|source| FontError::Allocation {
            resource: "prepared embedded fonts",
            source,
        })?;

    for (request, glyphs) in requests {
        let name = request.family().to_owned();
        let mut font = loader.load(&request)?;
        let properties = font
            .properties
            .ok_or_else(|| FontError::MissingProperties { name: name.clone() })?;
        validate_license(&name, properties.license())?;
        let subsetted = subset_requested && properties.license().may_subset();
        let data = if subsetted {
            let ids = glyph_ids(&font, &glyphs)?;
            subsetter.subset(&font, &ids)?
        } else {
            standalone(&mut font)?
        };
        prepared.push(Prepared {
            name,
            style: request.style(),
            data,
            properties,
            subsetted,
        });
    }
    Ok(prepared)
}

fn standalone(font: &mut FontData) -> Result<Vec<u8>, FontError> {
    let file = allsorts::binary::read::ReadScope::new(&font.data)
        .read::<allsorts::tables::OpenTypeFont<'_>>()
        .map_err(|error| FontError::EmbeddingFailed(format!("invalid OpenType font: {error}")))?;
    if !matches!(&file.data, allsorts::tables::OpenTypeData::Single(_)) || font.index != 0 {
        return Err(FontError::RequiresStandaloneFace);
    }
    Ok(std::mem::take(&mut font.data))
}

fn validate_license(name: &str, license: crate::License) -> Result<(), FontError> {
    if license.permission() == crate::Permission::Restricted {
        return Err(FontError::EmbeddingForbidden {
            name: name.to_owned(),
        });
    }
    if license
        .restrictions()
        .contains(crate::Restrictions::BITMAP_ONLY)
    {
        return Err(FontError::BitmapOnly {
            name: name.to_owned(),
        });
    }
    Ok(())
}
