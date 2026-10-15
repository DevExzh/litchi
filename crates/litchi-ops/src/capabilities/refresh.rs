//! Capability support for field refresh (records 0279 and 0295): DOCX and
//! DOC fields, DOCX tables of contents, PPTX and ODP slide fields, ODT
//! text fields, and workbook formula results.

#[cfg(any(
    feature = "docx",
    feature = "pptx",
    feature = "doc",
    feature = "odt",
    feature = "odp",
    all(
        feature = "eval",
        any(feature = "xlsx", feature = "xlsb", feature = "xls")
    )
))]
use super::RETAINED_SOURCE;
use super::{Capabilities, Refusal, Support};
use crate::Code;
#[cfg(any(
    feature = "docx",
    feature = "pptx",
    feature = "doc",
    feature = "odt",
    feature = "odp",
    all(
        feature = "eval",
        any(feature = "xlsx", feature = "xlsb", feature = "xls")
    )
))]
use crate::read::Adapter;

/// Refresh on a format without a refresh adapter.
const NOTHING_TO_REFRESH: Refusal = Refusal {
    code: Code::UnsupportedCapability,
    follow_up: "record 0295 follow-up: field refresh for PPT, RTF and the ODF object formats",
    reason: "this format has no field refresh adapter",
};

impl Capabilities {
    /// Whether [`super::Verb::Refresh`] is available.
    pub(super) fn refresh_support(&self) -> Support {
        if let Some(refusal) = self.read_refusal() {
            return refusal.support();
        }
        let (adapter, edits) = match self.adapter() {
            #[cfg(feature = "docx")]
            Some(Adapter::Docx) => (true, true),
            #[cfg(feature = "pptx")]
            Some(Adapter::Pptx) => (true, true),
            #[cfg(feature = "doc")]
            Some(Adapter::Doc) => (true, true),
            #[cfg(feature = "odt")]
            Some(Adapter::Odt) => (true, true),
            #[cfg(feature = "odp")]
            Some(Adapter::Odp) => (true, true),
            #[cfg(all(feature = "eval", feature = "xlsx"))]
            Some(Adapter::Xlsx) => (true, true),
            #[cfg(all(feature = "eval", feature = "xlsb"))]
            Some(Adapter::Xlsb) => (true, true),
            #[cfg(all(feature = "eval", feature = "xls"))]
            Some(Adapter::Xls) => (true, true),
            _ => (false, false),
        };
        if !adapter {
            return NOTHING_TO_REFRESH.support();
        }
        #[cfg(any(
            feature = "docx",
            feature = "pptx",
            feature = "doc",
            feature = "odt",
            feature = "odp",
            all(
                feature = "eval",
                any(feature = "xlsx", feature = "xlsb", feature = "xls")
            )
        ))]
        if edits && self.retains_source == Some(false) {
            return RETAINED_SOURCE.support();
        }
        let _ = edits;
        Support::Supported
    }
}
