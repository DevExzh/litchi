//! Format-independent font preparation and publication helpers.

pub mod powerpoint;

use crate::model::{FontData, FontError, FontProperties, GlyphMap, Request, Style};
#[cfg(feature = "subset")]
use crate::subset::mapping::glyph_ids_scoped;
#[cfg(feature = "subset")]
use crate::subset::{Allsorts, Limits as SubsetLimits, SubsetSession};

const DEFAULT_MAX_SUBSET_WORKING_BYTES: usize = 1024 * 1024 * 1024;

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

/// Resource bounds for resolving and preparing one document's font requests.
///
/// These limits apply before optional subsetting and again to the resulting
/// programs. They do not install, load, or render fonts.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PreparationLimits {
    pub max_requests: usize,
    pub max_glyphs_per_request: u64,
    pub max_family_name_bytes: usize,
    pub max_program_bytes: usize,
    pub max_total_program_bytes: usize,
    /// Maximum checked working-memory budget for one subset operation.
    ///
    /// This is independent from the retained output-program limits. The
    /// subsetter charges source, decoded structures, temporary indexes, and
    /// output assembly against this budget before the operation can proceed.
    pub max_subset_working_bytes: usize,
}

impl Default for PreparationLimits {
    fn default() -> Self {
        Self {
            max_requests: 1_024,
            max_glyphs_per_request: 1_114_112,
            max_family_name_bytes: 4_096,
            max_program_bytes: 64 * 1024 * 1024,
            max_total_program_bytes: 256 * 1024 * 1024,
            max_subset_working_bytes: DEFAULT_MAX_SUBSET_WORKING_BYTES,
        }
    }
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
/// The caller supplies request and program-size bounds explicitly.
pub fn prepare_with(
    resolver: &impl Resolver,
    used_glyphs: GlyphMap,
    mode: Mode,
    limits: PreparationLimits,
) -> Result<Vec<Prepared>, FontError> {
    #[cfg(not(feature = "subset"))]
    if mode == Mode::Subset {
        return Err(FontError::SubsettingUnavailable);
    }

    enforce_limit(
        "embedded-font requests",
        used_glyphs.len(),
        limits.max_requests,
    )?;

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

    let mut total_program_bytes = 0usize;
    for (request, glyphs) in requests {
        enforce_limit(
            "font-family name",
            request.family().len(),
            limits.max_family_name_bytes,
        )?;
        if glyphs.len() > limits.max_glyphs_per_request {
            return Err(FontError::LimitExceeded {
                resource: "glyphs in one font request",
                limit: usize::try_from(limits.max_glyphs_per_request).unwrap_or(usize::MAX),
                actual: usize::try_from(glyphs.len()).unwrap_or(usize::MAX),
            });
        }
        #[cfg(not(feature = "subset"))]
        let _ = &glyphs;
        let mut name = String::new();
        name.try_reserve_exact(request.family().len())
            .map_err(|source| FontError::Allocation {
                resource: "font-family name",
                source,
            })?;
        name.push_str(request.family());
        let mut font = resolver.resolve(&request)?;
        enforce_limit(
            "resolved font program",
            font.data.len(),
            limits.max_program_bytes,
        )?;
        let properties = font
            .properties
            .ok_or_else(|| FontError::MissingProperties { name: name.clone() })?;
        validate_license(&name, properties.license())?;
        let subsetted = mode == Mode::Subset && properties.license().may_subset();
        let data = if subsetted {
            #[cfg(feature = "subset")]
            {
                let remaining_total_program_bytes = limits
                    .max_total_program_bytes
                    .checked_sub(total_program_bytes)
                    .ok_or(FontError::LimitExceeded {
                        resource: "prepared font programs",
                        limit: limits.max_total_program_bytes,
                        actual: total_program_bytes,
                    })?;
                let output_limit = limits.max_program_bytes.min(remaining_total_program_bytes);
                let subset_session = SubsetSession::new(SubsetLimits::new(
                    limits.max_subset_working_bytes,
                    output_limit,
                ));
                subset_session.charge_mapping_staging(font.data.len(), glyphs.len())?;
                let ids = glyph_ids_scoped(&subset_session, &font, &glyphs)?;
                subsetter.subset_scoped(&subset_session, &font, &ids)?
            }
            #[cfg(not(feature = "subset"))]
            unreachable!("subset mode is rejected before resolution")
        } else {
            standalone(&mut font)?
        };
        enforce_limit(
            "prepared font program",
            data.len(),
            limits.max_program_bytes,
        )?;
        total_program_bytes =
            total_program_bytes
                .checked_add(data.len())
                .ok_or(FontError::LimitExceeded {
                    resource: "prepared font programs",
                    limit: limits.max_total_program_bytes,
                    actual: usize::MAX,
                })?;
        enforce_limit(
            "prepared font programs",
            total_program_bytes,
            limits.max_total_program_bytes,
        )?;
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

fn enforce_limit(resource: &'static str, actual: usize, limit: usize) -> Result<(), FontError> {
    if actual > limit {
        return Err(FontError::LimitExceeded {
            resource,
            limit,
            actual,
        });
    }
    Ok(())
}

/// Discover and prepare fonts using the host's configured system font source
/// under the caller-selected resource policy.
#[cfg(feature = "automatic")]
pub fn prepare(
    used_glyphs: GlyphMap,
    mode: Mode,
    limits: PreparationLimits,
) -> Result<Vec<Prepared>, FontError> {
    prepare_with(&automatic_loader(limits), used_glyphs, mode, limits)
}

#[cfg(feature = "automatic")]
fn automatic_loader(limits: PreparationLimits) -> crate::discovery::Loader {
    crate::discovery::Loader::with_max_font_program_bytes(
        limits.max_program_bytes.min(limits.max_total_program_bytes),
    )
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
    use std::mem::size_of;

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

    #[cfg(feature = "subset")]
    struct FixtureResolver;

    #[cfg(feature = "subset")]
    impl Resolver for FixtureResolver {
        fn resolve(&self, request: &Request) -> Result<FontData, FontError> {
            Ok(FontData {
                name: request.family().to_owned(),
                data: include_bytes!("../../../../test-data/fonts/NotoSansLycian-Regular.ttf")
                    .to_vec(),
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

    #[cfg(feature = "subset")]
    fn fixture_glyphs() -> GlyphMap {
        let mut glyphs = GlyphMap::new();
        glyphs.insert(
            Request::regular("Noto Sans Lycian"),
            ['\u{10280}'].into_iter().collect(),
        );
        glyphs
    }

    #[test]
    fn explicit_resolver_prepares_requests_in_stable_order() {
        let mut glyphs = HashMap::new();
        glyphs.insert(Request::regular("Zulu"), crate::Glyphs::new());
        glyphs.insert(Request::regular("Alpha"), crate::Glyphs::new());

        let prepared = prepare_with(
            &StaticResolver,
            glyphs,
            Mode::Full,
            PreparationLimits::default(),
        )
        .expect("prepared fonts");
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
            prepare_with(
                &StaticResolver,
                HashMap::new(),
                Mode::Subset,
                PreparationLimits::default(),
            ),
            Err(FontError::SubsettingUnavailable)
        ));
    }

    #[test]
    fn explicit_limits_bound_requests_before_resolution() {
        let mut glyphs = HashMap::new();
        glyphs.insert(Request::regular("Alpha"), crate::Glyphs::new());
        let limits = PreparationLimits {
            max_requests: 0,
            ..PreparationLimits::default()
        };
        assert!(matches!(
            prepare_with(&StaticResolver, glyphs, Mode::Full, limits),
            Err(FontError::LimitExceeded {
                resource: "embedded-font requests",
                ..
            })
        ));
    }

    #[test]
    fn preparation_limits_accept_exact_and_reject_over_boundaries() {
        fn glyph_map(family: &str, text: &str) -> GlyphMap {
            let mut map = GlyphMap::new();
            map.insert(Request::regular(family), text.chars().collect());
            map
        }

        let mut limits = PreparationLimits {
            max_family_name_bytes: 5,
            ..PreparationLimits::default()
        };
        assert!(
            prepare_with(&StaticResolver, glyph_map("Alpha", "A"), Mode::Full, limits,).is_ok()
        );
        limits.max_family_name_bytes = 4;
        assert!(matches!(
            prepare_with(&StaticResolver, glyph_map("Alpha", "A"), Mode::Full, limits,),
            Err(FontError::LimitExceeded {
                resource: "font-family name",
                limit: 4,
                actual: 5,
            })
        ));

        limits = PreparationLimits::default();
        limits.max_glyphs_per_request = 2;
        assert!(
            prepare_with(
                &StaticResolver,
                glyph_map("Alpha", "AB"),
                Mode::Full,
                limits,
            )
            .is_ok()
        );
        limits.max_glyphs_per_request = 1;
        assert!(matches!(
            prepare_with(
                &StaticResolver,
                glyph_map("Alpha", "AB"),
                Mode::Full,
                limits,
            ),
            Err(FontError::LimitExceeded {
                resource: "glyphs in one font request",
                limit: 1,
                actual: 2,
            })
        ));

        limits = PreparationLimits::default();
        limits.max_program_bytes = 12;
        assert!(
            prepare_with(&StaticResolver, glyph_map("Alpha", "A"), Mode::Full, limits,).is_ok()
        );
        limits.max_program_bytes = 11;
        assert!(matches!(
            prepare_with(&StaticResolver, glyph_map("Alpha", "A"), Mode::Full, limits,),
            Err(FontError::LimitExceeded {
                resource: "resolved font program",
                limit: 11,
                actual: 12,
            })
        ));

        let mut two = GlyphMap::new();
        two.insert(Request::regular("Alpha"), crate::Glyphs::new());
        two.insert(Request::regular("Zulu"), crate::Glyphs::new());
        limits = PreparationLimits::default();
        limits.max_total_program_bytes = 24;
        assert!(prepare_with(&StaticResolver, two.clone(), Mode::Full, limits).is_ok());
        limits.max_total_program_bytes = 23;
        assert!(matches!(
            prepare_with(&StaticResolver, two, Mode::Full, limits),
            Err(FontError::LimitExceeded {
                resource: "prepared font programs",
                limit: 23,
                actual: 24,
            })
        ));
    }

    #[cfg(feature = "subset")]
    #[test]
    fn subset_preparation_charges_staging_before_mapping() {
        assert_eq!(
            PreparationLimits::default().max_subset_working_bytes,
            DEFAULT_MAX_SUBSET_WORKING_BYTES
        );

        let source_bytes =
            include_bytes!("../../../../test-data/fonts/NotoSansLycian-Regular.ttf").len();
        let actual = source_bytes + (1 + 1) * size_of::<u16>();
        let limits = PreparationLimits {
            max_subset_working_bytes: actual - 1,
            ..PreparationLimits::default()
        };
        assert!(matches!(
            prepare_with(
                &FixtureResolver,
                fixture_glyphs(),
                Mode::Subset,
                limits,
            ),
            Err(FontError::LimitExceeded {
                resource: "font subsetting working bytes",
                limit,
                actual: observed,
            }) if limit + 1 == observed
        ));
    }

    #[cfg(feature = "subset")]
    #[test]
    fn subset_preparation_maps_output_budget_failures() {
        let limits = PreparationLimits {
            max_total_program_bytes: 1,
            ..PreparationLimits::default()
        };
        assert!(matches!(
            prepare_with(&FixtureResolver, fixture_glyphs(), Mode::Subset, limits),
            Err(FontError::LimitExceeded {
                resource: "font subsetting output bytes",
                limit: 1,
                actual,
            }) if actual > 1
        ));
    }

    #[cfg(feature = "subset")]
    #[test]
    fn subset_preparation_produces_a_real_fixture_program() {
        let prepared = prepare_with(
            &FixtureResolver,
            fixture_glyphs(),
            Mode::Subset,
            PreparationLimits::default(),
        )
        .expect("subset fixture");
        assert_eq!(prepared.len(), 1);
        assert!(prepared[0].subsetted);
        assert!(!prepared[0].data.is_empty());
    }

    #[cfg(feature = "automatic")]
    #[test]
    fn automatic_loader_uses_the_narrowest_selected_program_cap() {
        let limits = PreparationLimits {
            max_program_bytes: 17,
            max_total_program_bytes: 11,
            ..PreparationLimits::default()
        };
        assert_eq!(automatic_loader(limits).max_font_program_bytes(), 11);
    }
}
