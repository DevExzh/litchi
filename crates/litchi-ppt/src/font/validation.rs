use super::{Font, FontCollection, FontCollections, Limits};
use crate::package::{Error, Result};

pub(super) fn validate_limits(limits: Limits) -> Result<()> {
    if limits.max_fonts_per_collection > 129 {
        return Err(Error::Corrupted(
            "font count limit exceeds the format maximum".into(),
        ));
    }
    Ok(())
}

pub(super) fn validate_collections(value: &FontCollections, limits: Limits) -> Result<()> {
    validate_limits(limits)?;
    if let Some(base) = &value.base {
        validate_collection(base, limits)?;
    }
    if let Some(international) = &value.international {
        validate_collection(international, limits)?;
    }
    let (facets, bytes) = value
        .base
        .iter()
        .chain(value.international.iter())
        .flat_map(|collection| &collection.fonts)
        .flat_map(|font| &font.embedded_fonts)
        .try_fold((0usize, 0usize), |(facets, bytes), facet| {
            Ok::<_, Error>((
                facets
                    .checked_add(1)
                    .ok_or_else(|| Error::Corrupted("facet count overflow".into()))?,
                bytes
                    .checked_add(facet.data.len())
                    .ok_or_else(|| Error::Corrupted("embedded byte count overflow".into()))?,
            ))
        })?;
    if facets > limits.max_facets || bytes > limits.max_embedded_bytes {
        return Err(Error::Corrupted(
            "aggregate embedded fonts exceed configured limits".into(),
        ));
    }
    let font_count = value
        .base
        .iter()
        .chain(value.international.iter())
        .map(|collection| collection.fonts.len())
        .sum::<usize>();
    let copied = font_count
        .checked_mul(68)
        .and_then(|font_bytes| font_bytes.checked_add(bytes))
        .ok_or_else(|| Error::Corrupted("font copied-byte count overflow".into()))?;
    let record_count = font_count
        .checked_add(facets)
        .and_then(|count| count.checked_add(value.base.is_some() as usize))
        .and_then(|count| count.checked_add(value.international.is_some() as usize))
        .and_then(|count| count.checked_add(value.embedding_flags.is_some() as usize))
        .ok_or_else(|| Error::Corrupted("font record count overflow".into()))?;
    if copied > limits.records.max_copied_payload_bytes
        || record_count > limits.records.max_records
        || (record_count != 0 && limits.records.max_depth == 0)
    {
        return Err(Error::ResourceLimit(
            "font owners exceed the composed record budget".into(),
        ));
    }
    if value.base.as_ref().is_some_and(|v| v.international)
        || value
            .international
            .as_ref()
            .is_some_and(|v| !v.international)
    {
        return Err(Error::Corrupted(
            "font collection is stored in the wrong namespace".into(),
        ));
    }
    if let Some(flags) = value.embedding_flags
        && (flags.subset != (flags.raw & 1 != 0)
            || flags.subset_option_confirmed != (flags.raw & 2 != 0))
    {
        return Err(Error::Corrupted(
            "embedding flag projections disagree with raw bits".into(),
        ));
    }
    Ok(())
}

pub(super) fn validate_collection(value: &FontCollection, limits: Limits) -> Result<()> {
    validate_limits(limits)?;
    if value.fonts.len() > limits.max_fonts_per_collection {
        return Err(Error::Corrupted("font collection exceeds its limit".into()));
    }
    let mut facets = 0usize;
    let mut bytes = 0usize;
    for (ordinal, font) in value.fonts.iter().enumerate() {
        validate_font(font)?;
        if usize::from(font.index) != ordinal {
            return Err(Error::Corrupted(
                "font ordinals are not stable and contiguous".into(),
            ));
        }
        for (position, facet) in font.embedded_fonts.iter().enumerate() {
            if facet.style > 3
                || position > 0 && font.embedded_fonts[position - 1].style >= facet.style
            {
                return Err(Error::Corrupted(
                    "font facets are duplicated or out of order".into(),
                ));
            }
            facets = facets
                .checked_add(1)
                .ok_or_else(|| Error::Corrupted("facet count overflow".into()))?;
            bytes = bytes
                .checked_add(facet.data.len())
                .ok_or_else(|| Error::Corrupted("embedded byte count overflow".into()))?;
            if facet.data.len() > limits.max_facet_bytes {
                return Err(Error::Corrupted(
                    "embedded facet exceeds its byte limit".into(),
                ));
            }
        }
    }
    if facets > limits.max_facets || bytes > limits.max_embedded_bytes {
        return Err(Error::Corrupted(
            "embedded fonts exceed configured limits".into(),
        ));
    }
    Ok(())
}

pub(super) fn validate_authored_collection(value: &FontCollection, limits: Limits) -> Result<()> {
    validate_collection(value, limits)?;
    for font in &value.fonts {
        for facet in &font.embedded_fonts {
            super::validate_eot_facet(facet.data.as_ref(), limits)?;
        }
    }
    Ok(())
}

pub(super) fn validate_font(font: &Font) -> Result<()> {
    if font.raw_instance > 128 {
        return Err(Error::Corrupted(
            "FontEntityAtom instance exceeds 128".into(),
        ));
    }
    let units = font.name.encode_utf16().count();
    if units > 32 || font.name.chars().any(|c| c == '\0') {
        return Err(Error::Corrupted(
            "font name exceeds the 32 UTF-16 unit field or contains NUL".into(),
        ));
    }
    if font.embedded_subset != (font.font_flags & 1 != 0)
        || font.raster != (font.font_type_flags & 1 != 0)
        || font.device != (font.font_type_flags & 2 != 0)
        || font.truetype != (font.font_type_flags & 4 != 0)
        || font.no_substitution != (font.font_type_flags & 8 != 0)
    {
        return Err(Error::Corrupted(
            "font projections disagree with raw flag bits".into(),
        ));
    }
    Ok(())
}
