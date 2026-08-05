//! Typed OLE-control metadata owned by the `parts::ole::controls` context.

/// The OLE controls recorded in a document.
///
/// The records are inert metadata. A `Controls` value never instantiates or
/// activates an OLE control and never executes control code.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Controls {
    pub(crate) controls: Vec<Control>,
}

/// One `OcxInfo` entry (MS-DOC 2.9.161).
///
/// MS-DOC defines the cookie as a unique index within the document's
/// `RgxOcxInfo` table. The remaining bytes of a padded entry are deliberately
/// not interpreted.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Control {
    /// Unique index of this control within the document's `RgxOcxInfo`.
    pub cookie: u32,
}

impl Controls {
    pub(crate) fn from_controls(controls: Vec<Control>) -> Self {
        Self { controls }
    }

    /// All recorded OLE controls, in table order.
    pub fn controls(&self) -> &[Control] {
        &self.controls
    }

    /// Number of recorded OLE controls.
    pub fn len(&self) -> usize {
        self.controls.len()
    }

    /// Whether the table contains no OLE controls.
    pub fn is_empty(&self) -> bool {
        self.controls.is_empty()
    }
}
