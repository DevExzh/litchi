/// Hyperlink structure for Word documents
use super::parts::hyperlinks::{Hyperlink as InternalHyperlink, HyperlinkType};

/// A hyperlink in a Word document
#[derive(Debug, Clone)]
pub struct Hyperlink {
    /// Character position range
    pub start_position: u32,
    pub end_position: u32,
    /// Destination URL/path
    pub destination: String,
    /// Display text
    pub display_text: String,
    /// Link type
    pub link_type: HyperlinkType,
}

impl Hyperlink {
    /// Create from internal hyperlink
    pub(crate) fn from_internal(internal: &InternalHyperlink) -> Self {
        Self {
            start_position: internal.start_cp,
            end_position: internal.end_cp,
            destination: internal.destination.clone(),
            display_text: internal.display_text.clone(),
            link_type: internal.link_type.clone(),
        }
    }

    /// Get the destination
    pub fn destination(&self) -> &str {
        &self.destination
    }

    /// Get the display text
    pub fn display_text(&self) -> &str {
        &self.display_text
    }

    /// Check if this is a URL hyperlink
    pub fn is_url(&self) -> bool {
        self.link_type == HyperlinkType::Url
    }

    /// Check if this is an email hyperlink
    pub fn is_email(&self) -> bool {
        self.link_type == HyperlinkType::Email
    }

    /// Check if this is a file hyperlink
    pub fn is_file(&self) -> bool {
        self.link_type == HyperlinkType::File
    }

    /// Check if this is a bookmark hyperlink
    pub fn is_bookmark(&self) -> bool {
        self.link_type == HyperlinkType::Bookmark
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_hyperlink_from_internal_url() {
        let internal = InternalHyperlink {
            start_cp: 100,
            end_cp: 200,
            field_code: "HYPERLINK \"http://example.com\"".to_string(),
            destination: "http://example.com".to_string(),
            display_text: "Example Site".to_string(),
            link_type: HyperlinkType::Url,
        };

        let link = Hyperlink::from_internal(&internal);

        assert_eq!(link.start_position, 100);
        assert_eq!(link.end_position, 200);
        assert_eq!(link.destination, "http://example.com");
        assert_eq!(link.display_text, "Example Site");
        assert_eq!(link.link_type, HyperlinkType::Url);
    }

    #[test]
    fn test_hyperlink_from_internal_email() {
        let internal = InternalHyperlink {
            start_cp: 50,
            end_cp: 150,
            field_code: "HYPERLINK \"mailto:test@example.com\"".to_string(),
            destination: "test@example.com".to_string(),
            display_text: "Email Us".to_string(),
            link_type: HyperlinkType::Email,
        };

        let link = Hyperlink::from_internal(&internal);

        assert_eq!(link.start_position, 50);
        assert_eq!(link.end_position, 150);
        assert_eq!(link.destination(), "test@example.com");
        assert_eq!(link.display_text(), "Email Us");
        assert!(link.is_email());
        assert!(!link.is_url());
        assert!(!link.is_file());
        assert!(!link.is_bookmark());
    }

    #[test]
    fn test_hyperlink_from_internal_file() {
        let internal = InternalHyperlink {
            start_cp: 200,
            end_cp: 300,
            field_code: "HYPERLINK \"C:\\\\path\\\\file.txt\"".to_string(),
            destination: "C:\\\\path\\\\file.txt".to_string(),
            display_text: "Open File".to_string(),
            link_type: HyperlinkType::File,
        };

        let link = Hyperlink::from_internal(&internal);

        assert!(link.is_file());
        assert!(!link.is_url());
        assert!(!link.is_email());
        assert!(!link.is_bookmark());
    }

    #[test]
    fn test_hyperlink_from_internal_bookmark() {
        let internal = InternalHyperlink {
            start_cp: 300,
            end_cp: 400,
            field_code: "HYPERLINK \\\\l \"Section1\"".to_string(),
            destination: "#Section1".to_string(),
            display_text: "Go to Section".to_string(),
            link_type: HyperlinkType::Bookmark,
        };

        let link = Hyperlink::from_internal(&internal);

        assert!(link.is_bookmark());
        assert!(!link.is_url());
        assert!(!link.is_email());
        assert!(!link.is_file());
    }

    #[test]
    fn test_hyperlink_destination_method() {
        let link = Hyperlink::from_internal(&InternalHyperlink {
            start_cp: 0,
            end_cp: 100,
            field_code: String::new(),
            destination: "https://example.com".to_string(),
            display_text: String::new(),
            link_type: HyperlinkType::Url,
        });

        assert_eq!(link.destination(), "https://example.com");
    }

    #[test]
    fn test_hyperlink_display_text_method() {
        let link = Hyperlink::from_internal(&InternalHyperlink {
            start_cp: 0,
            end_cp: 100,
            field_code: String::new(),
            destination: String::new(),
            display_text: "Click Here".to_string(),
            link_type: HyperlinkType::Url,
        });

        assert_eq!(link.display_text(), "Click Here");
    }

    #[test]
    fn test_hyperlink_clone() {
        let internal = InternalHyperlink {
            start_cp: 100,
            end_cp: 200,
            field_code: "HYPERLINK \"http://example.com\"".to_string(),
            destination: "http://example.com".to_string(),
            display_text: "Example".to_string(),
            link_type: HyperlinkType::Url,
        };

        let link = Hyperlink::from_internal(&internal);
        let cloned = link.clone();

        assert_eq!(cloned.start_position, link.start_position);
        assert_eq!(cloned.end_position, link.end_position);
        assert_eq!(cloned.destination, link.destination);
        assert_eq!(cloned.display_text, link.display_text);
        assert_eq!(cloned.link_type, link.link_type);
    }

    #[test]
    fn test_hyperlink_debug() {
        let internal = InternalHyperlink {
            start_cp: 100,
            end_cp: 200,
            field_code: "HYPERLINK \"http://example.com\"".to_string(),
            destination: "http://example.com".to_string(),
            display_text: "Debug Test".to_string(),
            link_type: HyperlinkType::Url,
        };

        let link = Hyperlink::from_internal(&internal);
        let debug_str = format!("{:?}", link);

        assert!(debug_str.contains("Hyperlink"));
        assert!(debug_str.contains("http://example.com"));
        assert!(debug_str.contains("Debug Test"));
        assert!(debug_str.contains("100"));
        assert!(debug_str.contains("200"));
    }

    #[test]
    fn test_hyperlink_link_type_equality() {
        let link_url = Hyperlink::from_internal(&InternalHyperlink {
            start_cp: 0,
            end_cp: 100,
            field_code: String::new(),
            destination: "http://example.com".to_string(),
            display_text: "URL".to_string(),
            link_type: HyperlinkType::Url,
        });

        let link_email = Hyperlink::from_internal(&InternalHyperlink {
            start_cp: 0,
            end_cp: 100,
            field_code: String::new(),
            destination: "mailto:test@example.com".to_string(),
            display_text: "Email".to_string(),
            link_type: HyperlinkType::Email,
        });

        assert_ne!(link_url.link_type, link_email.link_type);
        assert!(link_url.is_url());
        assert!(link_email.is_email());
    }

    #[test]
    fn test_hyperlink_empty_fields() {
        let link = Hyperlink::from_internal(&InternalHyperlink {
            start_cp: 0,
            end_cp: 0,
            field_code: String::new(),
            destination: String::new(),
            display_text: String::new(),
            link_type: HyperlinkType::Unknown,
        });

        assert_eq!(link.destination(), "");
        assert_eq!(link.display_text(), "");
        assert!(!link.is_url());
        assert!(!link.is_email());
        assert!(!link.is_file());
        assert!(!link.is_bookmark());
    }

    #[test]
    fn test_hyperlink_unicode() {
        let link = Hyperlink::from_internal(&InternalHyperlink {
            start_cp: 100,
            end_cp: 200,
            field_code: "HYPERLINK \"https://例子.com\"".to_string(),
            destination: "https://例子.com".to_string(),
            display_text: "Unicode: 你好世界 🎉".to_string(),
            link_type: HyperlinkType::Url,
        });

        assert_eq!(link.destination(), "https://例子.com");
        assert_eq!(link.display_text(), "Unicode: 你好世界 🎉");
    }
}
