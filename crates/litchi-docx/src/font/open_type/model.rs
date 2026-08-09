#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
#![expect(
    clippy::unwrap_used,
    reason = "the invariant is established immediately before extraction"
)]
//! Semantic values for the Word 2010 OpenType `rPr` extension family.

use crate::error::{Error, Result};
use std::fmt;
use std::num::NonZeroU8;

/// Word 2010 OpenType extension namespace.
pub(super) const WORD_2010_NAMESPACE: &[u8] =
    b"http://schemas.microsoft.com/office/word/2010/wordml";

/// Markup-compatibility namespace used by `mc:Ignorable`.
pub(super) const MC_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/markup-compatibility/2006";

/// Ligature features defined by `ST_Ligatures`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Ligatures {
    None,
    Standard,
    Contextual,
    Historical,
    Discretional,
    StandardContextual,
    StandardHistorical,
    ContextualHistorical,
    StandardDiscretional,
    ContextualDiscretional,
    HistoricalDiscretional,
    StandardContextualHistorical,
    StandardContextualDiscretional,
    StandardHistoricalDiscretional,
    ContextualHistoricalDiscretional,
    All,
}

impl Ligatures {
    pub(crate) fn parse(value: &str) -> Result<Self> {
        let value = match value {
            "none" => Self::None,
            "standard" => Self::Standard,
            "contextual" => Self::Contextual,
            "historical" => Self::Historical,
            "discretional" => Self::Discretional,
            "standardContextual" => Self::StandardContextual,
            "standardHistorical" => Self::StandardHistorical,
            "contextualHistorical" => Self::ContextualHistorical,
            "standardDiscretional" => Self::StandardDiscretional,
            "contextualDiscretional" => Self::ContextualDiscretional,
            "historicalDiscretional" => Self::HistoricalDiscretional,
            "standardContextualHistorical" => Self::StandardContextualHistorical,
            "standardContextualDiscretional" => Self::StandardContextualDiscretional,
            "standardHistoricalDiscretional" => Self::StandardHistoricalDiscretional,
            "contextualHistoricalDiscretional" => Self::ContextualHistoricalDiscretional,
            "all" => Self::All,
            _ => {
                return Err(invalid(format!(
                    "invalid OpenType ligatures value '{value}'"
                )));
            },
        };
        Ok(value)
    }

    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::None => "none",
            Self::Standard => "standard",
            Self::Contextual => "contextual",
            Self::Historical => "historical",
            Self::Discretional => "discretional",
            Self::StandardContextual => "standardContextual",
            Self::StandardHistorical => "standardHistorical",
            Self::ContextualHistorical => "contextualHistorical",
            Self::StandardDiscretional => "standardDiscretional",
            Self::ContextualDiscretional => "contextualDiscretional",
            Self::HistoricalDiscretional => "historicalDiscretional",
            Self::StandardContextualHistorical => "standardContextualHistorical",
            Self::StandardContextualDiscretional => "standardContextualDiscretional",
            Self::StandardHistoricalDiscretional => "standardHistoricalDiscretional",
            Self::ContextualHistoricalDiscretional => "contextualHistoricalDiscretional",
            Self::All => "all",
        }
    }
}

impl fmt::Display for Ligatures {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

/// Numeral form defined by `ST_NumForm`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum NumForm {
    Default,
    Lining,
    OldStyle,
}

impl NumForm {
    pub(crate) fn parse(value: &str) -> Result<Self> {
        match value {
            "default" => Ok(Self::Default),
            "lining" => Ok(Self::Lining),
            "oldStyle" => Ok(Self::OldStyle),
            _ => Err(invalid(format!("invalid OpenType numForm value '{value}'"))),
        }
    }

    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Default => "default",
            Self::Lining => "lining",
            Self::OldStyle => "oldStyle",
        }
    }
}

/// Numeral spacing defined by `ST_NumSpacing`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum NumSpacing {
    Default,
    Proportional,
    Tabular,
}

impl NumSpacing {
    pub(crate) fn parse(value: &str) -> Result<Self> {
        match value {
            "default" => Ok(Self::Default),
            "proportional" => Ok(Self::Proportional),
            "tabular" => Ok(Self::Tabular),
            _ => Err(invalid(format!(
                "invalid OpenType numSpacing value '{value}'"
            ))),
        }
    }

    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Default => "default",
            Self::Proportional => "proportional",
            Self::Tabular => "tabular",
        }
    }
}

/// The authored state of a `CT_OnOff` element.
///
/// `None` means that the element was present without `val`; this is distinct
/// from an absent element (`Option<OnOff>::None`) and from explicit `true` or
/// `false`.  The wire default for an authored element without `val` is true.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct OnOff {
    authored: Option<bool>,
}

impl OnOff {
    /// Construct an authored `CT_OnOff`; `None` emits an empty/default-on
    /// element, while `Some` emits the corresponding `ST_OnOff` value.
    #[must_use]
    pub const fn new(authored: Option<bool>) -> Self {
        Self { authored }
    }

    /// Construct a present empty element whose schema default is true.
    #[must_use]
    pub const fn default_on() -> Self {
        Self::new(None)
    }

    /// Construct an explicit `val="true"` element.
    #[must_use]
    pub const fn on() -> Self {
        Self::new(Some(true))
    }

    /// Construct an explicit `val="false"` element.
    #[must_use]
    pub const fn off() -> Self {
        Self::new(Some(false))
    }

    /// Return the authored `val`, preserving omission as `None`.
    #[must_use]
    pub const fn authored(self) -> Option<bool> {
        self.authored
    }

    /// Return the effective schema value for an authored element.
    #[must_use]
    pub const fn effective(self) -> bool {
        match self.authored {
            Some(value) => value,
            None => true,
        }
    }

    /// Alias for [`Self::effective`] using the concise semantic vocabulary of
    /// a checked boolean value.
    #[must_use]
    pub const fn value(self) -> bool {
        self.effective()
    }
}

/// Checked identifier for a Word stylistic set (`1..=20`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[repr(transparent)]
pub struct StyleSetId(NonZeroU8);

impl StyleSetId {
    /// Construct a stylistic-set identifier in the schema range.
    #[must_use]
    ///
    /// # Panics
    ///
    /// Panics if an internal writer invariant is violated.
    pub const fn new(value: u8) -> Option<Self> {
        if value == 0 || value > 20 {
            None
        } else {
            Some(Self(NonZeroU8::new(value).unwrap()))
        }
    }

    /// Return the wire identifier.
    #[must_use]
    pub const fn get(self) -> u8 {
        self.0.get()
    }
}

impl TryFrom<u8> for StyleSetId {
    type Error = Error;

    fn try_from(value: u8) -> Result<Self> {
        Self::new(value)
            .ok_or_else(|| invalid(format!("stylistic set id {value} is outside 1..=20")))
    }
}

/// One `w14:styleSet` child.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct StyleSet {
    pub id: StyleSetId,
    /// Authored `w14:val`; omission means enabled by the schema default.
    pub enabled: Option<bool>,
}

impl StyleSet {
    #[must_use]
    pub const fn new(id: StyleSetId) -> Self {
        Self { id, enabled: None }
    }

    #[must_use]
    pub const fn with_enabled(mut self, enabled: Option<bool>) -> Self {
        self.enabled = enabled;
        self
    }
}

/// The complete typed Word 2010 OpenType extension family on one run.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct OpenType {
    pub ligatures: Option<Ligatures>,
    pub num_form: Option<NumForm>,
    pub num_spacing: Option<NumSpacing>,
    pub stylistic_sets: Vec<StyleSet>,
    pub cntxt_alts: Option<OnOff>,
    stylistic_sets_present: bool,
}

impl OpenType {
    /// Parse a complete `w:r` or `w:rPr` fragment.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn parse(xml: &[u8]) -> Result<Self> {
        super::codec::parse(xml)
    }

    /// Validate all semantic domains and duplicate identifiers.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn validate(&self) -> Result<()> {
        super::validation::validate(self)
    }

    /// Replace one stylistic set by id, or append it when absent.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn set_style_set(&mut self, value: StyleSet) -> Result<&mut Self> {
        super::validation::validate_style_set(&value)?;
        self.stylistic_sets_present = true;
        if let Some(existing) = self
            .stylistic_sets
            .iter_mut()
            .find(|existing| existing.id == value.id)
        {
            *existing = value;
        } else {
            self.stylistic_sets.push(value);
        }
        self.validate()?;
        Ok(self)
    }

    /// Remove one stylistic set by its checked identifier.
    pub fn remove_style_set(&mut self, id: StyleSetId) -> Option<StyleSet> {
        self.stylistic_sets
            .iter()
            .position(|value| value.id == id)
            .map(|index| self.stylistic_sets.remove(index))
    }

    /// Whether the `w14:stylisticSets` container was authored.
    ///
    /// This remains distinct from an authored empty container and from an
    /// absent container, so a source-preserving edit can retain both forms.
    #[must_use]
    pub const fn stylistic_sets_present(&self) -> bool {
        self.stylistic_sets_present
    }

    /// Set or remove the `w14:stylisticSets` container while retaining its
    /// ordered style-set children.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn set_stylistic_sets(&mut self, value: Option<Vec<StyleSet>>) -> Result<&mut Self> {
        if let Some(values) = value {
            if values.len() > super::validation::MAX_STYLE_SETS {
                return Err(invalid("too many OpenType stylistic sets"));
            }
            self.stylistic_sets = values;
            self.stylistic_sets_present = true;
        } else {
            self.stylistic_sets.clear();
            self.stylistic_sets_present = false;
        }
        self.validate()?;
        Ok(self)
    }

    /// Remove the authored stylistic-set container and all of its children.
    pub fn clear_stylistic_sets(&mut self) -> &mut Self {
        self.stylistic_sets.clear();
        self.stylistic_sets_present = false;
        self
    }

    /// Find one stylistic set by id.
    #[must_use]
    pub fn style_set(&self, id: StyleSetId) -> Option<StyleSet> {
        self.stylistic_sets
            .iter()
            .copied()
            .find(|value| value.id == id)
    }
}

pub(crate) fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
