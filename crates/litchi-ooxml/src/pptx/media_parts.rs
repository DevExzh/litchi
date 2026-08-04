//! Compatibility exports for the canonical PPTX media-parts codec.
//!
//! Semantic media values, bounded XML handling, inert resource loading, and
//! package mutation live in litchi_pptx::media_parts. These wrappers keep
//! the historical host import path and OoxmlError result boundary.

use crate::error::{OoxmlError, Result};
use litchi_opc::{OpcPackage, PackURI};

pub use litchi_pptx::media_parts::{
    MediaBookmark, MediaData, MediaExtensionList, MediaFade, MediaResource, MediaTrim,
    OfficeMediaExtension, SlideMediaConformance, SlideMediaKind, SlideMediaList, SlideMediaPicture,
    SlideMediaPoster, SlideMediaTransform,
};

fn map_media_error(error: litchi_pptx::Error) -> OoxmlError {
    match error {
        litchi_pptx::Error::Opc(error) => OoxmlError::Opc(error),
        litchi_pptx::Error::Xml(message) => OoxmlError::Xml(message),
        litchi_pptx::Error::Invalid(message) => OoxmlError::InvalidFormat(message),
        litchi_pptx::Error::Limit { resource, limit }
            if resource == "slide media serialized XML bytes" =>
        {
            OoxmlError::Pptx(litchi_pptx::Error::Limit { resource, limit })
        },
        litchi_pptx::Error::Limit { resource, .. } => {
            OoxmlError::InvalidFormat(format!("{resource} limit exceeded"))
        },
        litchi_pptx::Error::Allocation { resource, source } => {
            OoxmlError::Allocation { resource, source }
        },
        litchi_pptx::Error::MarkupCompatibility(error) => {
            OoxmlError::Common(litchi_ooxml_common::Error::Mce(error))
        },
        litchi_pptx::Error::Decode(error) => {
            OoxmlError::Common(litchi_ooxml_common::Error::Decode(error))
        },
        other => OoxmlError::Pptx(other),
    }
}

/// Parse all audio/video pictures from a complete Slide part.
pub fn parse_slide_media(xml: &[u8]) -> Result<SlideMediaList> {
    litchi_pptx::media_parts::parse_slide_media(xml).map_err(map_media_error)
}
/// Serialize audio/video pictures for insertion into a Slide part.
pub fn write_slide_media_pictures(
    value: &SlideMediaList,
    conformance: SlideMediaConformance,
) -> Result<Vec<u8>> {
    litchi_pptx::media_parts::write_slide_media_pictures(value, conformance)
        .map_err(map_media_error)
}

/// Load media pictures and validate their complete internal OPC resource graph.
pub fn load_slide_media(package: &OpcPackage, slide_name: &PackURI) -> Result<SlideMediaList> {
    litchi_pptx::media_parts::load_slide_media(package, slide_name).map_err(map_media_error)
}

/// Add media pictures and their inert internal resources to a Slide part.
pub fn store_slide_media(
    package: &mut OpcPackage,
    slide_name: &PackURI,
    value: &SlideMediaList,
    conformance: SlideMediaConformance,
) -> Result<()> {
    litchi_pptx::media_parts::store_slide_media(package, slide_name, value, conformance)
        .map_err(map_media_error)
}

#[cfg(test)]
mod tests {
    use super::*;

    const POI_VIDEO: &[u8] =
        include_bytes!("../../../../test-data/poi/test-data/slideshow/EmbeddedVideo.pptx");

    #[test]
    fn historical_host_shim_loads_media_fixture() {
        let package = OpcPackage::from_bytes(POI_VIDEO).expect("POI video package");
        let slide = PackURI::new("/ppt/slides/slide1.xml").expect("slide URI");
        let media = load_slide_media(&package, &slide).expect("media graph");

        assert_eq!(media.pictures.len(), 1);
        assert_eq!(media.pictures[0].kind, SlideMediaKind::Video);
        assert_eq!(
            media.pictures[0].resource.as_ref().unwrap().data.len(),
            101_799
        );
    }

    #[test]
    fn historical_host_shim_keeps_owner_failures_at_the_ooxml_boundary() {
        let error = parse_slide_media(b"<!DOCTYPE invalid>").expect_err("invalid XML");
        assert!(matches!(error, OoxmlError::InvalidFormat(_)));
    }
}
