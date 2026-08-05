//! Host error-boundary adapter for standalone PPTX controls.

use crate::error::Result;
use crate::pptx::media_parts::map_pptx_error;
use litchi_opc::{OpcPackage, Part};

pub use litchi_pptx::presentation::embedded::controls::{
    Binary as ControlBinary, Control as SlideControl, Descriptor as ControlDescriptor, Persistence,
};

/// Host-internal name retained for callers that share discovery bounds.
pub(crate) type ControlLoadLimits = litchi_pptx::presentation::embedded::controls::Limits;

/// Discover controls through the standalone owner and translate only its
/// failure boundary into the legacy host error type.
pub(crate) fn load_slide_controls(
    package: &OpcPackage,
    slide_index: usize,
    slide: &dyn Part,
    limits: &mut ControlLoadLimits,
) -> Result<Vec<SlideControl>> {
    litchi_pptx::presentation::embedded::controls::load_slide(package, slide_index, slide, limits)
        .map_err(map_pptx_error)
}
