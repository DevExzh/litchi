//! Capability support for template merge (records 0274 and 0295): the
//! OOXML, legacy, RTF, ODT/OTH/ODM, ODS and ODP templates.

#[cfg(any(
    feature = "docx",
    feature = "pptx",
    feature = "xlsx",
    feature = "doc",
    feature = "ppt",
    feature = "xls",
    feature = "xlsb",
    feature = "rtf",
    feature = "odt",
    feature = "ods",
    feature = "odp"
))]
use super::RETAINED_SOURCE;
use super::{Capabilities, Refusal, Support};
use crate::Code;
#[cfg(any(
    feature = "docx",
    feature = "pptx",
    feature = "xlsx",
    feature = "doc",
    feature = "ppt",
    feature = "xls",
    feature = "xlsb",
    feature = "rtf",
    feature = "odt",
    feature = "ods",
    feature = "odp"
))]
use crate::read::Adapter;

/// Merge on a format without a merge adapter.
const NOT_A_TEMPLATE: Refusal = Refusal {
    code: Code::UnsupportedCapability,
    follow_up: "template merge for the ODF object formats and iWork (record 0295)",
    reason: "template merge fills word-processing, presentation and spreadsheet templates",
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
            #[cfg(feature = "doc")]
            Some(Adapter::Doc) => true,
            #[cfg(feature = "ppt")]
            Some(Adapter::Ppt) => true,
            #[cfg(feature = "xls")]
            Some(Adapter::Xls) => true,
            #[cfg(feature = "xlsb")]
            Some(Adapter::Xlsb) => true,
            #[cfg(feature = "rtf")]
            Some(Adapter::Rtf) => true,
            #[cfg(feature = "odt")]
            Some(Adapter::Odt) => true,
            #[cfg(feature = "ods")]
            Some(Adapter::Ods) => true,
            #[cfg(feature = "odp")]
            Some(Adapter::Odp) => true,
            _ => false,
        };
        if !template {
            return NOT_A_TEMPLATE.support();
        }
        #[cfg(any(
            feature = "docx",
            feature = "pptx",
            feature = "xlsx",
            feature = "doc",
            feature = "ppt",
            feature = "xls",
            feature = "xlsb",
            feature = "rtf",
            feature = "odt",
            feature = "ods",
            feature = "odp"
        ))]
        if self.retains_source == Some(false) {
            return RETAINED_SOURCE.support();
        }
        Support::Supported
    }
}
