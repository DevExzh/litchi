//! Semantic hyperlink target values.

use crate::Result;

/// A hyperlink attached to a shape or text run.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Hyperlink {
    /// An external URL or an unrecognized inert target scheme.
    External {
        url: String,
        tooltip: Option<String>,
    },
    /// A one-based slide number in the current presentation.
    Slide {
        slide_number: usize,
        tooltip: Option<String>,
    },
    /// An email target with an optional subject.
    Email {
        email: String,
        subject: Option<String>,
        tooltip: Option<String>,
    },
}

impl Hyperlink {
    /// Construct an external URL target.
    pub fn url(url: impl Into<String>) -> Self {
        Self::External {
            url: url.into(),
            tooltip: None,
        }
    }

    /// Construct an external URL target with a tooltip.
    pub fn url_with_tooltip(url: impl Into<String>, tooltip: impl Into<String>) -> Self {
        Self::External {
            url: url.into(),
            tooltip: Some(tooltip.into()),
        }
    }

    /// Construct a one-based internal slide target.
    #[must_use]
    pub fn slide(slide_number: usize) -> Self {
        Self::Slide {
            slide_number,
            tooltip: None,
        }
    }

    /// Construct a slide target with a tooltip.
    pub fn slide_with_tooltip(slide_number: usize, tooltip: impl Into<String>) -> Self {
        Self::Slide {
            slide_number,
            tooltip: Some(tooltip.into()),
        }
    }

    /// Construct an email target.
    pub fn email(email: impl Into<String>) -> Self {
        Self::Email {
            email: email.into(),
            subject: None,
            tooltip: None,
        }
    }

    /// Construct an email target with an optional subject and tooltip.
    pub fn email_with_subject(
        email: impl Into<String>,
        subject: impl Into<String>,
        tooltip: Option<String>,
    ) -> Self {
        Self::Email {
            email: email.into(),
            subject: Some(subject.into()),
            tooltip,
        }
    }

    /// Parse an inert target and optional tooltip.
    ///
    /// The parser is kept on the value facade for discoverability while the
    /// bounded grammar lives in the sibling codec module.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_xml(target: &str, tooltip: Option<String>) -> Result<Self> {
        super::codec::parse(target, tooltip)
    }

    /// Return the target spelling used by a writer.
    #[must_use]
    pub fn target(&self) -> String {
        match self {
            Self::External { url, .. } => url.clone(),
            Self::Slide { slide_number, .. } => {
                format!("ppaction://hlinksldjump?sldNum={slide_number}")
            },
            Self::Email { email, subject, .. } => subject.as_ref().map_or_else(
                || format!("mailto:{email}"),
                |subject| format!("mailto:{email}?subject={subject}"),
            ),
        }
    }

    /// Return the optional tooltip without copying it.
    #[must_use]
    pub fn tooltip(&self) -> Option<&str> {
        match self {
            Self::External { tooltip, .. }
            | Self::Slide { tooltip, .. }
            | Self::Email { tooltip, .. } => tooltip.as_deref(),
        }
    }

    /// Return whether this target leaves the current presentation.
    #[must_use]
    pub fn is_external(&self) -> bool {
        matches!(self, Self::External { .. } | Self::Email { .. })
    }
}
