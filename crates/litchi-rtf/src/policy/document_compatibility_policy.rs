/// Requested UI feature-throttling behavior.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DocumentFeatureThrottle {
    /// Limit incompatible UI functionality (`nofeaturethrottle0`).
    CompatibilityLimited,
    /// Do not limit incompatible UI functionality (`nofeaturethrottle1`).
    Unrestricted,
}

impl DocumentFeatureThrottle {
    pub(crate) fn from_rtf(value: i32) -> Option<Self> {
        match value {
            0 => Some(Self::CompatibilityLimited),
            1 => Some(Self::Unrestricted),
            _ => None,
        }
    }

    pub(crate) const fn rtf_value(self) -> i32 {
        match self {
            Self::CompatibilityLimited => 0,
            Self::Unrestricted => 1,
        }
    }
}

/// Passive document compatibility reset, UI, and upgrade requests.
///
/// This crate retains these declarations but does not reset compatibility
/// options, disable UI, or upgrade document contents.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct DocumentCompatibilityPolicy {
    /// `\nocompatoptions`: reset all compatibility options to defaults.
    pub reset_options_to_defaults: bool,
    /// Explicit `\nofeaturethrottleN` policy; `\nouicompat` maps to unrestricted.
    pub feature_throttle: Option<DocumentFeatureThrottle>,
    /// `\forceupgrade`: request an application-defined document upgrade.
    pub force_upgrade: bool,
}

impl DocumentCompatibilityPolicy {
    /// Return whether every compatibility policy declaration was omitted.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        !self.reset_options_to_defaults && self.feature_throttle.is_none() && !self.force_upgrade
    }

    /// Return the explicit policy or the specification's omission behavior.
    #[must_use]
    pub fn effective_feature_throttle(&self) -> DocumentFeatureThrottle {
        self.feature_throttle
            .unwrap_or(DocumentFeatureThrottle::CompatibilityLimited)
    }
}
