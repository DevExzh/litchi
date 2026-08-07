//! Format-independent font preparation and publication helpers.

pub mod powerpoint;

use crate::model::{FontData, FontError, FontProperties, GlyphMap, Request, Style};
#[cfg(feature = "subset")]
use crate::subset::{Allsorts, Subsetter, glyph_ids};

/// Resolves one typed face request to an owned font program and metadata.
pub trait Resolver {
    fn resolve(&self, request: &Request) -> Result<FontData, FontError>;
}

/// Whether preparation keeps a complete face or creates a glyph subset.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum Mode {
    #[default]
    Full,
    Subset,
}

/// One discovered and optionally subsetted, but still un-obfuscated, program.
#[derive(Debug)]
pub struct Prepared {
    pub name: String,
    pub style: Style,
    pub data: Vec<u8>,
    pub properties: FontProperties,
    pub subsetted: bool,
}

/// Resolve and prepare fonts without knowing any package format.
///
/// The returned buffers are owned and un-obfuscated. Publication, content
/// types, relationship IDs, and part names belong to the document adapters.
/// Requests are processed in key order, making resolver calls and results
/// deterministic regardless of `GlyphMap`'s hash iteration order.
pub fn prepare_with(
    resolver: &impl Resolver,
    used_glyphs: GlyphMap,
    mode: Mode,
) -> Result<Vec<Prepared>, FontError> {
    #[cfg(not(feature = "subset"))]
    if mode == Mode::Subset {
        return Err(FontError::SubsettingUnavailable);
    }

    #[cfg(feature = "subset")]
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
        #[cfg(not(feature = "subset"))]
        let _ = &glyphs;
        let name = request.family().to_owned();
        let mut font = resolver.resolve(&request)?;
        let properties = font
            .properties
            .ok_or_else(|| FontError::MissingProperties { name: name.clone() })?;
        validate_license(&name, properties.license())?;
        let subsetted = mode == Mode::Subset && properties.license().may_subset();
        let data = if subsetted {
            #[cfg(feature = "subset")]
            {
                let ids = glyph_ids(&font, &glyphs)?;
                subsetter.subset(&font, &ids)?
            }
            #[cfg(not(feature = "subset"))]
            unreachable!("subset mode is rejected before resolution")
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

/// Discover and prepare fonts using the host's configured system font source.
#[cfg(feature = "automatic")]
pub fn prepare(used_glyphs: GlyphMap, mode: Mode) -> Result<Vec<Prepared>, FontError> {
    prepare_with(&crate::discovery::Loader::new(), used_glyphs, mode)
}

fn standalone(font: &mut FontData) -> Result<Vec<u8>, FontError> {
    let signature = font.data.get(..4).ok_or_else(|| {
        FontError::EmbeddingFailed("invalid OpenType font: missing sfnt signature".into())
    })?;
    if signature == b"ttcf" || font.index != 0 {
        return Err(FontError::RequiresStandaloneFace);
    }
    if font.data.len() < 12 || !matches!(signature, b"\0\x01\0\0" | b"OTTO" | b"true" | b"typ1") {
        return Err(FontError::EmbeddingFailed(
            "invalid OpenType font: unsupported sfnt signature".into(),
        ));
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

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::HashMap;

    struct StaticResolver;

    impl Resolver for StaticResolver {
        fn resolve(&self, request: &Request) -> Result<FontData, FontError> {
            Ok(FontData {
                name: request.family().to_owned(),
                data: vec![0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                index: 0,
                properties: Some(FontProperties::new(
                    crate::License::new(0).expect("installable license"),
                    crate::Panose::new([0; 10]),
                    None,
                    crate::Family::Auto,
                    crate::Pitch::Default,
                    crate::Signature::new([0; 4], [0; 2]),
                )),
            })
        }
    }

    #[test]
    fn explicit_resolver_prepares_requests_in_stable_order() {
        let mut glyphs = HashMap::new();
        glyphs.insert(Request::regular("Zulu"), crate::Glyphs::new());
        glyphs.insert(Request::regular("Alpha"), crate::Glyphs::new());

        let prepared = prepare_with(&StaticResolver, glyphs, Mode::Full).expect("prepared fonts");
        assert_eq!(
            prepared
                .iter()
                .map(|font| font.name.as_str())
                .collect::<Vec<_>>(),
            ["Alpha", "Zulu"]
        );
    }

    #[cfg(not(feature = "subset"))]
    #[test]
    fn subset_mode_requires_the_subset_capability() {
        assert!(matches!(
            prepare_with(&StaticResolver, HashMap::new(), Mode::Subset),
            Err(FontError::SubsettingUnavailable)
        ));
    }
}
