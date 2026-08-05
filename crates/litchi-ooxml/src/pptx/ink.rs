//! Host error-boundary adapter for standalone PPTX InkML resources.

use crate::error::Result;
use crate::pptx::media_parts::map_pptx_error;
use litchi_opc::{OpcPackage, PackURI, Part};

pub use litchi_pptx::presentation::embedded::ink::Annotation;

/// The OPC content type of an InkML part.
pub const INK_CONTENT_TYPE: &str = litchi_pptx::presentation::embedded::ink::CONTENT_TYPE;

/// Host-internal name retained for shared discovery bounds.
pub(crate) type InkLoadLimits = litchi_pptx::presentation::embedded::ink::Limits;

/// Host spelling retained for the standalone store result.
pub type StoredInkAnnotation = litchi_pptx::presentation::embedded::ink::StoredAnnotation;

/// Discover InkML content parts through the standalone owner.
pub(crate) fn load_slide_ink_annotations(
    package: &OpcPackage,
    slide_index: usize,
    slide: &dyn Part,
    limits: &mut InkLoadLimits,
) -> Result<Vec<Annotation>> {
    litchi_pptx::presentation::embedded::ink::load_slide(package, slide_index, slide, limits)
        .map_err(map_pptx_error)
}

/// Store an InkML content part through the standalone owner.
pub fn store_slide_ink_annotation(
    package: &mut OpcPackage,
    slide_name: &PackURI,
    inkml: &[u8],
) -> Result<StoredInkAnnotation> {
    litchi_pptx::presentation::embedded::ink::store_slide(package, slide_name, inkml)
        .map_err(map_pptx_error)
}
