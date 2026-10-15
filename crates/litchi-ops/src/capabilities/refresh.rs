//! Capability support for field refresh (record 0279): DOCX fields and
//! tables of contents, PPTX slide fields, and workbook formula results.

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
    follow_up: "field refresh for the legacy, RTF and ODF formats (record 0279 follow-up)",
    reason: "field refresh recomputes DOCX fields, PPTX slide fields and workbook formulas",
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
