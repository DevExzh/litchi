//! Capability support for text find/replace (record 0273): `set` with the
//! `find=` and `replace=` properties.

#[cfg(any(feature = "docx", feature = "xlsx", feature = "pptx"))]
use super::RETAINED_SOURCE;
use super::{Capabilities, Refusal, Support};
use crate::Code;
use crate::read::Adapter;

/// The follow-up of every format whose adapter has no find yet.
const FOLLOW_UP: &str =
    "record 0273 follow-up: a find adapter over the format's run-level text edits";

impl Capabilities {
    /// Whether `set` with `find=` is available.
    pub(super) fn replace_support(&self) -> Support {
        if let Some(refusal) = self.read_refusal() {
            return refusal.support();
        }
        let reason = match self.adapter() {
            #[cfg(feature = "docx")]
            Some(Adapter::Docx) => return self.retained(),
            #[cfg(feature = "xlsx")]
            Some(Adapter::Xlsx) => return self.retained(),
            #[cfg(feature = "pptx")]
            Some(Adapter::Pptx) => return self.retained(),
            #[cfg(feature = "doc")]
            Some(Adapter::Doc) => {
                "DOC find/replace is not integrated; the leaf's body run spans can carry it"
            },
            #[cfg(feature = "ppt")]
            Some(Adapter::Ppt) => {
                "PPT find/replace is not integrated; the leaf replaces shape text whole"
            },
            #[cfg(feature = "xls")]
            Some(Adapter::Xls) => "XLS find/replace in cell text is not integrated",
            #[cfg(feature = "xlsb")]
            Some(Adapter::Xlsb) => "XLSB find/replace in cell text is not integrated",
            #[cfg(feature = "rtf")]
            Some(Adapter::Rtf) => {
                "RTF find/replace is not integrated; the leaf's run text edits can carry it"
            },
            _ => "find/replace is not integrated for this format",
        };
        Refusal {
            code: Code::UnsupportedCapability,
            follow_up: FOLLOW_UP,
            reason,
        }
        .support()
    }

    #[cfg(any(feature = "docx", feature = "xlsx", feature = "pptx"))]
    fn retained(&self) -> Support {
        if self.retains_source == Some(false) {
            return RETAINED_SOURCE.support();
        }
        Support::Supported
    }
}
