//! Compatibility facade for the canonical PPTX action-settings owner.
//!
//! Semantic values and bounded package discovery live in
//! [`litchi_pptx::actions`]. This adapter retains the historical OOXML names
//! and maps relationship failures back to their established `OoxmlError`
//! variants.

use crate::error::{OoxmlError, Result};
use litchi_opc::{OpcPackage, Part};

pub use litchi_pptx::actions::{
    PptxActionKind, PptxActionSetting, PptxActionTarget, PptxActionTrigger, PptxSlideShowJump,
};

pub(crate) use litchi_pptx::actions::Limits as ActionLoadLimits;

pub(crate) fn load_slide_action_settings(
    package: &OpcPackage,
    slide_index: usize,
    slide: &dyn Part,
    limits: &mut ActionLoadLimits,
) -> Result<Vec<PptxActionSetting>> {
    litchi_pptx::actions::load_slide_action_settings(package, slide_index, slide, limits)
        .map_err(map_error)
}

fn map_error(error: litchi_pptx::Error) -> OoxmlError {
    match error {
        litchi_pptx::Error::Opc(error) => OoxmlError::Opc(error),
        litchi_pptx::Error::Xml(message) => OoxmlError::Xml(message),
        litchi_pptx::Error::Invalid(message) => OoxmlError::InvalidFormat(message),
        litchi_pptx::Error::Limit { resource, .. } => {
            OoxmlError::InvalidFormat(format!("{resource} exceeds the supported safety limit"))
        },
        litchi_pptx::Error::Allocation { resource, source } => {
            OoxmlError::Allocation { resource, source }
        },
        litchi_pptx::Error::MarkupCompatibility(error) => OoxmlError::from(error),
        litchi_pptx::Error::Decode(error) => OoxmlError::from(error),
        litchi_pptx::Error::Relationship(message) => OoxmlError::InvalidRelationship(message),
        litchi_pptx::Error::PartNotFound(message) => OoxmlError::PartNotFound(message),
        error => OoxmlError::Pptx(error),
    }
}
