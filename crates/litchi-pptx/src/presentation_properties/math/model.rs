//! Package-independent values for the `PresentationML` math extension.

use crate::{Error, Result};

/// Placement of a binary operator when an equation wraps.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum BinaryBreak {
    /// Place the operator before the line break.
    Before,
    /// Place the operator after the line break.
    After,
    /// Repeat the operator on both sides of the line break.
    Repeat,
}

impl BinaryBreak {
    pub(crate) const fn wire_value(self) -> &'static str {
        match self {
            Self::Before => "before",
            Self::After => "after",
            Self::Repeat => "repeat",
        }
    }

    pub(crate) fn from_wire(value: &str) -> Result<Self> {
        match value {
            "before" => Ok(Self::Before),
            "after" => Ok(Self::After),
            "repeat" => Ok(Self::Repeat),
            _ => Err(invalid(format!("invalid brkBin value '{value}'"))),
        }
    }
}

/// Placement of a repeated subtraction operator when equations wrap.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum BinarySubtractionBreak {
    /// Repeat a minus sign on both lines.
    MinusMinus,
    /// Use plus on the first line and minus on the second.
    PlusMinus,
    /// Use minus on the first line and plus on the second.
    MinusPlus,
}

impl BinarySubtractionBreak {
    pub(crate) const fn wire_value(self) -> &'static str {
        match self {
            Self::MinusMinus => "--",
            Self::PlusMinus => "+-",
            Self::MinusPlus => "-+",
        }
    }

    pub(crate) fn from_wire(value: &str) -> Result<Self> {
        match value {
            "--" => Ok(Self::MinusMinus),
            "+-" => Ok(Self::PlusMinus),
            "-+" => Ok(Self::MinusPlus),
            _ => Err(invalid(format!("invalid brkBinSub value '{value}'"))),
        }
    }
}

/// Bounded document-level math defaults stored in a14:m.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Properties {
    /// Optional m:brkBin setting.
    pub binary_break: Option<BinaryBreak>,
    /// Optional m:brkBinSub setting.
    pub binary_subtraction_break: Option<BinarySubtractionBreak>,
}

impl Properties {
    /// Construct an empty math-properties snapshot.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            binary_break: None,
            binary_subtraction_break: None,
        }
    }

    /// Set the binary-operator wrapping policy.
    #[must_use]
    pub const fn with_binary_break(mut self, value: BinaryBreak) -> Self {
        self.binary_break = Some(value);
        self
    }

    /// Set the binary-subtraction wrapping policy.
    #[must_use]
    pub const fn with_binary_subtraction_break(mut self, value: BinarySubtractionBreak) -> Self {
        self.binary_subtraction_break = Some(value);
        self
    }

    /// Validate the package-independent math snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn validate(&self) -> Result<()> {
        // The enums are closed over the schema lexical domains. Keeping this
        // method explicit gives future revisions a validation seam without
        // making the XML codec responsible for semantic snapshot checks.
        if self.binary_break.is_none() && self.binary_subtraction_break.is_none() {
            return Ok(());
        }
        Ok(())
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
