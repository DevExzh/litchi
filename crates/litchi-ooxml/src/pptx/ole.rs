//! Host error-boundary adapter for standalone PPTX OLE resources.

use crate::error::Result;
use crate::pptx::media_parts::map_pptx_error;
use litchi_opc::{OpcPackage, Part};

pub use litchi_pptx::presentation::embedded::ole::{Kind as PayloadKind, Mode, Object, Target};

/// Host-internal name retained for shared OLE discovery bounds.
pub(crate) type OleLoadLimits = litchi_pptx::presentation::embedded::ole::Limits;

/// Discover OLE shapes through the standalone owner.
pub(crate) fn load_slide_ole_objects(
    package: &OpcPackage,
    slide_index: usize,
    slide: &dyn Part,
    limits: &mut OleLoadLimits,
) -> Result<Vec<Object>> {
    litchi_pptx::presentation::embedded::ole::load_slide(package, slide_index, slide, limits)
        .map_err(map_pptx_error)
}
