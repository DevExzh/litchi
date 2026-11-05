//! Capability support for the M4 verbs of record 0266: CSV/TSV import and
//! formula evaluation, both worksheet operations of XLSX, XLSB and XLS; ODS
//! imports since record 0295.

#[cfg(any(feature = "xlsx", feature = "xlsb", feature = "xls", feature = "ods"))]
use super::RETAINED_SOURCE;
use super::{Capabilities, Refusal, Support, Verb};
use crate::Code;
#[cfg(any(feature = "xlsx", feature = "xlsb", feature = "xls", feature = "ods"))]
use crate::read::Adapter;

/// Import or evaluation on a format without worksheets.
const NOT_A_WORKBOOK: Refusal = Refusal {
    code: Code::UnsupportedCapability,
    follow_up: "none: import and evaluation are worksheet operations",
    reason: "this format has no worksheets",
};

/// Evaluation of an OpenDocument spreadsheet (records 0277 and 0295).
#[cfg(feature = "ods")]
const ODS_PENDING: Refusal = Refusal {
    code: Code::UnsupportedCapability,
    follow_up: "record 0295 follow-up: OpenFormula translation for the litchi-eval engine",
    reason: "ODS formulas are OpenFormula; the evaluation engine reads Excel syntax",
};

/// Evaluation in a build without the formula engine.
#[cfg(not(feature = "eval"))]
const NO_ENGINE: Refusal = Refusal {
    code: Code::UnsupportedCapability,
    follow_up: "build litchi-ops with the `eval` feature",
    reason: "formula evaluation needs the litchi-eval engine, which this build leaves out",
};

impl Capabilities {
    /// Whether `verb` ([`Verb::Import`] or [`Verb::Evaluate`]) is available.
    pub(super) fn m4_support(&self, verb: Verb) -> Support {
        if let Some(refusal) = self.read_refusal() {
            return refusal.support();
        }
        #[cfg(feature = "ods")]
        if self.adapter() == Some(Adapter::Ods) {
            return match verb {
                Verb::Import if self.retains_source == Some(false) => RETAINED_SOURCE.support(),
                Verb::Import => Support::Supported,
                _ => ODS_PENDING.support(),
            };
        }
        let workbook = match self.adapter() {
            #[cfg(feature = "xlsx")]
            Some(Adapter::Xlsx) => true,
            #[cfg(feature = "xlsb")]
            Some(Adapter::Xlsb) => true,
            #[cfg(feature = "xls")]
            Some(Adapter::Xls) => true,
            _ => false,
        };
        if !workbook {
            return NOT_A_WORKBOOK.support();
        }
        match verb {
            #[cfg(any(feature = "xlsx", feature = "xlsb", feature = "xls"))]
            Verb::Import if self.retains_source == Some(false) => RETAINED_SOURCE.support(),
            #[cfg(not(feature = "eval"))]
            Verb::Evaluate => NO_ENGINE.support(),
            _ => Support::Supported,
        }
    }
}
