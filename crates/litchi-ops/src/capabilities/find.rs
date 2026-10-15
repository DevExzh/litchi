//! Capability support for text find/replace (record 0273): `set` with the
//! `find=` and `replace=` properties.

#[cfg(any(
    feature = "docx",
    feature = "xlsx",
    feature = "pptx",
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
use crate::read::Adapter;

/// The follow-up of every format whose adapter has no find yet (the ODF
/// object formats and iWork).
const FOLLOW_UP: &str = "record 0295 follow-up: a find adapter over the format's text edits";

impl Capabilities {
    /// Whether `set` with `find=` is available.
    pub(super) fn replace_support(&self) -> Support {
        if let Some(refusal) = self.read_refusal() {
            return refusal.support();
        }
        match self.adapter() {
            #[cfg(feature = "docx")]
            Some(Adapter::Docx) => self.retained(),
            #[cfg(feature = "xlsx")]
            Some(Adapter::Xlsx) => self.retained(),
            #[cfg(feature = "pptx")]
            Some(Adapter::Pptx) => self.retained(),
            // Record 0295: the text-surface adapters.
            #[cfg(feature = "doc")]
            Some(Adapter::Doc) => self.retained(),
            #[cfg(feature = "ppt")]
            Some(Adapter::Ppt) => self.retained(),
            #[cfg(feature = "xls")]
            Some(Adapter::Xls) => self.retained(),
            #[cfg(feature = "xlsb")]
            Some(Adapter::Xlsb) => self.retained(),
            #[cfg(feature = "rtf")]
            Some(Adapter::Rtf) => self.retained(),
            #[cfg(feature = "odt")]
            Some(Adapter::Odt) => self.retained(),
            #[cfg(feature = "ods")]
            Some(Adapter::Ods) => self.retained(),
            #[cfg(feature = "odp")]
            Some(Adapter::Odp) => self.retained(),
            _ => Refusal {
                code: Code::UnsupportedCapability,
                follow_up: FOLLOW_UP,
                reason: "find/replace is not integrated for this format",
            }
            .support(),
        }
    }

    #[cfg(any(
        feature = "docx",
        feature = "xlsx",
        feature = "pptx",
        feature = "doc",
        feature = "ppt",
        feature = "xls",
        feature = "xlsb",
        feature = "rtf",
        feature = "odt",
        feature = "ods",
        feature = "odp"
    ))]
    fn retained(&self) -> Support {
        if self.retains_source == Some(false) {
            return RETAINED_SOURCE.support();
        }
        Support::Supported
    }
}
