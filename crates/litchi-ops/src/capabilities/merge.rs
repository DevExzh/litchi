//! Capability support for template merge (record 0274): DOCX, PPTX and
//! XLSX templates.

#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx"))]
use super::RETAINED_SOURCE;
use super::{Capabilities, Refusal, Support};
use crate::Code;
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx"))]
use crate::read::Adapter;

/// Merge on a format without a merge adapter.
const NOT_A_TEMPLATE: Refusal = Refusal {
    code: Code::UnsupportedCapability,
    follow_up: "template merge for XLSB and the legacy, RTF and ODF formats (record 0274 F1)",
    reason: "template merge fills DOCX, PPTX and XLSX templates",
};

impl Capabilities {
    /// Whether [`super::Verb::Merge`] is available.
    pub(super) fn merge_support(&self) -> Support {
        if let Some(refusal) = self.read_refusal() {
            return refusal.support();
        }
        let template = match self.adapter() {
            #[cfg(feature = "docx")]
            Some(Adapter::Docx) => true,
            #[cfg(feature = "pptx")]
            Some(Adapter::Pptx) => true,
            #[cfg(feature = "xlsx")]
            Some(Adapter::Xlsx) => true,
            _ => false,
        };
        if !template {
            return NOT_A_TEMPLATE.support();
        }
        #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx"))]
        if self.retains_source == Some(false) {
            return RETAINED_SOURCE.support();
        }
        Support::Supported
    }
}
