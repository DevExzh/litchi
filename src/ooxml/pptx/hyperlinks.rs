//! Hyperlink support for PowerPoint presentations.
//!
//! This module provides types for working with hyperlinks in slides.

use crate::ooxml::error::Result;

/// A hyperlink in a presentation.
///
/// Can link to external URLs, other slides, or email addresses.
#[derive(Debug, Clone, PartialEq)]
pub enum Hyperlink {
    /// External URL hyperlink
    External {
        /// URL to link to
        url: String,
        /// Optional tooltip text
        tooltip: Option<String>,
    },
    /// Internal slide hyperlink
    Slide {
        /// Slide number to link to (1-based)
        slide_number: usize,
        /// Optional tooltip text
        tooltip: Option<String>,
    },
    /// Email hyperlink
    Email {
        /// Email address
        email: String,
        /// Optional subject
        subject: Option<String>,
        /// Optional tooltip text
        tooltip: Option<String>,
    },
}

#[allow(dead_code)] // Part of the public API for future use
impl Hyperlink {
    /// Create an external URL hyperlink.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::ooxml::pptx::Hyperlink;
    ///
    /// let link = Hyperlink::url("https://example.com");
    /// ```
    pub fn url(url: impl Into<String>) -> Self {
        Hyperlink::External {
            url: url.into(),
            tooltip: None,
        }
    }

    /// Create an external URL hyperlink with tooltip.
    pub fn url_with_tooltip(url: impl Into<String>, tooltip: impl Into<String>) -> Self {
        Hyperlink::External {
            url: url.into(),
            tooltip: Some(tooltip.into()),
        }
    }

    /// Create a slide hyperlink.
    ///
    /// # Arguments
    /// * `slide_number` - 1-based slide number to link to
    pub fn slide(slide_number: usize) -> Self {
        Hyperlink::Slide {
            slide_number,
            tooltip: None,
        }
    }

    /// Create a slide hyperlink with tooltip.
    pub fn slide_with_tooltip(slide_number: usize, tooltip: impl Into<String>) -> Self {
        Hyperlink::Slide {
            slide_number,
            tooltip: Some(tooltip.into()),
        }
    }

    /// Create an email hyperlink.
    pub fn email(email: impl Into<String>) -> Self {
        Hyperlink::Email {
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
        Hyperlink::Email {
            email: email.into(),
            subject: Some(subject.into()),
            tooltip,
        }
    }

    /// Get the target for this hyperlink (for XML generation).
    pub(crate) fn target(&self) -> String {
        match self {
            Hyperlink::External { url, .. } => url.clone(),
            Hyperlink::Slide { slide_number, .. } => {
                format!("ppaction://hlinksldjump?sldNum={}", slide_number)
            },
            Hyperlink::Email { email, subject, .. } => {
                if let Some(subj) = subject {
                    format!("mailto:{}?subject={}", email, subj)
                } else {
                    format!("mailto:{}", email)
                }
            },
        }
    }

    /// Get the tooltip if present.
    pub fn tooltip(&self) -> Option<&str> {
        match self {
            Hyperlink::External { tooltip, .. } => tooltip.as_deref(),
            Hyperlink::Slide { tooltip, .. } => tooltip.as_deref(),
            Hyperlink::Email { tooltip, .. } => tooltip.as_deref(),
        }
    }

    /// Check if this is an external hyperlink.
    pub fn is_external(&self) -> bool {
        matches!(self, Hyperlink::External { .. } | Hyperlink::Email { .. })
    }

    /// Parse hyperlink from XML attributes and relationship.
    pub(crate) fn from_xml(target: &str, tooltip: Option<String>) -> Result<Self> {
        if target.starts_with("http://") || target.starts_with("https://") {
            Ok(Hyperlink::External {
                url: target.to_string(),
                tooltip,
            })
        } else if target.starts_with("ppaction://hlinksldjump") {
            // Extract slide number from ppaction URL
            let slide_number = target
                .split("sldNum=")
                .nth(1)
                .and_then(|s| s.parse::<usize>().ok())
                .unwrap_or(1);
            Ok(Hyperlink::Slide {
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
                    .find(|p| p.starts_with("subject="))
                    .map(|s| s.trim_start_matches("subject=").to_string())
            } else {
                None
            };
            Ok(Hyperlink::Email {
                email,
                subject,
                tooltip,
            })
        } else {
            // Default to external
            Ok(Hyperlink::External {
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
    fn test_external_hyperlink() {
        let link = Hyperlink::url("https://example.com");
        assert!(link.is_external());
        assert_eq!(link.target(), "https://example.com");
    }

    #[test]
    fn test_slide_hyperlink() {
        let link = Hyperlink::slide(3);
        assert!(!link.is_external());
        assert_eq!(link.target(), "ppaction://hlinksldjump?sldNum=3");
    }

    #[test]
    fn test_email_hyperlink() {
        let link = Hyperlink::email("test@example.com");
        assert!(link.is_external());
        assert_eq!(link.target(), "mailto:test@example.com");
    }

    #[test]
    fn test_email_with_subject() {
        let link = Hyperlink::email_with_subject("test@example.com", "Hello", None);
        assert_eq!(link.target(), "mailto:test@example.com?subject=Hello");
    }
}
