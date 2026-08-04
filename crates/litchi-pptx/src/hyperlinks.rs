//! Inert hyperlink values used by PresentationML.
//!
//! The package host resolves relationship targets and owns XML traversal. This
//! module only models hyperlink values and translates the target strings used
//! by PowerPoint; it never follows a URL or accesses package storage.

use crate::Result;

/// A hyperlink in a presentation.
///
/// A hyperlink can point to an external URL, another slide, or an email
/// address. Tooltip text is retained when it is present in the source markup.
#[derive(Debug, Clone, PartialEq)]
pub enum Hyperlink {
    /// External URL hyperlink.
    External {
        /// URL to link to.
        url: String,
        /// Optional tooltip text.
        tooltip: Option<String>,
    },
    /// Internal slide hyperlink.
    Slide {
        /// One-based slide number to link to.
        slide_number: usize,
        /// Optional tooltip text.
        tooltip: Option<String>,
    },
    /// Email hyperlink.
    Email {
        /// Email address.
        email: String,
        /// Optional subject.
        subject: Option<String>,
        /// Optional tooltip text.
        tooltip: Option<String>,
    },
}

#[allow(dead_code)] // `target` is used by the future PresentationML writer.
impl Hyperlink {
    /// Create an external URL hyperlink.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi_pptx::Hyperlink;
    ///
    /// let link = Hyperlink::url("https://example.com");
    /// ```
    pub fn url(url: impl Into<String>) -> Self {
        Self::External {
            url: url.into(),
            tooltip: None,
        }
    }

    /// Create an external URL hyperlink with tooltip.
    pub fn url_with_tooltip(url: impl Into<String>, tooltip: impl Into<String>) -> Self {
        Self::External {
            url: url.into(),
            tooltip: Some(tooltip.into()),
        }
    }

    /// Create a slide hyperlink.
    ///
    /// # Arguments
    /// * `slide_number` - One-based slide number to link to.
    pub fn slide(slide_number: usize) -> Self {
        Self::Slide {
            slide_number,
            tooltip: None,
        }
    }

    /// Create a slide hyperlink with tooltip.
    pub fn slide_with_tooltip(slide_number: usize, tooltip: impl Into<String>) -> Self {
        Self::Slide {
            slide_number,
            tooltip: Some(tooltip.into()),
        }
    }

    /// Create an email hyperlink.
    pub fn email(email: impl Into<String>) -> Self {
        Self::Email {
            email: email.into(),
            subject: None,
            tooltip: None,
        }
    }

    /// Create an email hyperlink with subject and tooltip.
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

    /// Get the target string used when serializing this hyperlink.
    pub(crate) fn target(&self) -> String {
        match self {
            Self::External { url, .. } => url.clone(),
            Self::Slide { slide_number, .. } => {
                format!("ppaction://hlinksldjump?sldNum={slide_number}")
            },
            Self::Email { email, subject, .. } => {
                if let Some(subject) = subject {
                    format!("mailto:{email}?subject={subject}")
                } else {
                    format!("mailto:{email}")
                }
            },
        }
    }

    /// Get the tooltip if present.
    pub fn tooltip(&self) -> Option<&str> {
        match self {
            Self::External { tooltip, .. }
            | Self::Slide { tooltip, .. }
            | Self::Email { tooltip, .. } => tooltip.as_deref(),
        }
    }

    /// Check if this is an external hyperlink, including an email hyperlink.
    pub fn is_external(&self) -> bool {
        matches!(self, Self::External { .. } | Self::Email { .. })
    }

    /// Parse an inert PowerPoint hyperlink target and optional tooltip.
    ///
    /// The OOXML host calls this after resolving a relationship or reading an
    /// inline `a:hlinkClick` action. This function does not access a package,
    /// resolve relationships, or follow the target. Unknown target schemes
    /// retain the historical behavior of being represented as external URLs.
    pub fn from_xml(target: &str, tooltip: Option<String>) -> Result<Self> {
        if target.starts_with("http://") || target.starts_with("https://") {
            Ok(Self::External {
                url: target.to_string(),
                tooltip,
            })
        } else if target.starts_with("ppaction://hlinksldjump") {
            // Extract slide number from the PowerPoint action target.
            let slide_number = target
                .split("sldNum=")
                .nth(1)
                .and_then(|value| value.parse::<usize>().ok())
                .unwrap_or(1);
            Ok(Self::Slide {
                slide_number,
                tooltip,
            })
        } else if target.starts_with("mailto:") {
            let email_part = target.trim_start_matches("mailto:");
            let parts: Vec<&str> = email_part.split('?').collect();
            let email = parts[0].to_string();
            let subject = if parts.len() > 1 {
                parts[1]
                    .split('&')
                    .find(|part| part.starts_with("subject="))
                    .map(|part| part.trim_start_matches("subject=").to_string())
            } else {
                None
            };
            Ok(Self::Email {
                email,
                subject,
                tooltip,
            })
        } else {
            // Preserve unknown target schemes as external targets.
            Ok(Self::External {
                url: target.to_string(),
                tooltip,
            })
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn external_hyperlink_preserves_target_and_tooltip() {
        let link = Hyperlink::url_with_tooltip("https://example.com", "Open site");
        assert!(link.is_external());
        assert_eq!(link.tooltip(), Some("Open site"));
        assert_eq!(link.target(), "https://example.com");
    }

    #[test]
    fn slide_hyperlink_uses_powerpoint_slide_jump_target() {
        let link = Hyperlink::slide(3);
        assert!(!link.is_external());
        assert_eq!(link.target(), "ppaction://hlinksldjump?sldNum=3");
    }

    #[test]
    fn email_hyperlink_preserves_subject_target_behavior() {
        let link = Hyperlink::email_with_subject("test@example.com", "Hello", None);
        assert!(link.is_external());
        assert_eq!(link.target(), "mailto:test@example.com?subject=Hello");
    }

    #[test]
    fn parser_preserves_powerpoint_targets_and_tooltips() {
        let slide = Hyperlink::from_xml(
            "ppaction://hlinksldjump?sldNum=7",
            Some("Next section".to_string()),
        )
        .unwrap();
        assert_eq!(slide, Hyperlink::slide_with_tooltip(7, "Next section"));

        let email = Hyperlink::from_xml(
            "mailto:test@example.com?cc=other@example.com&subject=Hello",
            None,
        )
        .unwrap();
        assert_eq!(
            email,
            Hyperlink::Email {
                email: "test@example.com".to_string(),
                subject: Some("Hello".to_string()),
                tooltip: None,
            }
        );
    }
}
