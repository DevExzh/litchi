//! Typed value for the inert data-type-icon visibility extensions.

use super::invalid;

/// The `visible` value of one MS-XLSX data-type-icon extension.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct ShowDataTypeIcons {
    visible: bool,
}

impl ShowDataTypeIcons {
    /// Construct a value using the `CT_ShowDataTypeIcons/@visible` boolean.
    #[must_use]
    pub const fn new(visible: bool) -> Self {
        Self { visible }
    }

    /// The effective visibility.  An omitted wire attribute is `true`.
    #[must_use]
    pub const fn visible(self) -> bool {
        self.visible
    }

    pub(crate) fn parse_boolean(value: &str) -> crate::Result<bool> {
        // XML Schema `boolean` has `whiteSpace=collapse`, whose whitespace
        // class is exactly XML S (#x9, #xA, #xD, and #x20). Rust's general
        // `str::trim` also accepts Unicode whitespace such as NBSP, which is
        // outside that value space.
        match value.trim_matches(|character| matches!(character, ' ' | '\t' | '\r' | '\n')) {
            "true" | "1" => Ok(true),
            "false" | "0" => Ok(false),
            _ => Err(invalid(format!(
                "invalid {} boolean '{}'; expected true, false, 1, or 0",
                "showDataTypeIcons", value
            ))),
        }
    }
}

impl Default for ShowDataTypeIcons {
    fn default() -> Self {
        Self::new(true)
    }
}
