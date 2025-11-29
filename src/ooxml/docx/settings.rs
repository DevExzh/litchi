/// Document settings and protection support.
///
/// This module provides types and methods for accessing document settings
/// and protection status.
use crate::ooxml::error::{OoxmlError, Result};
use crate::ooxml::opc::part::Part;
use quick_xml::Reader;
use quick_xml::events::Event;

/// Document settings including protection status.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi::ooxml::docx::Package;
///
/// let pkg = Package::open("document.docx")?;
/// let doc = pkg.document()?;
///
/// if let Some(settings) = doc.settings()? {
///     if settings.is_protected() {
///         println!("Document is protected");
///         println!("Protection type: {:?}", settings.protection_type());
///     }
/// }
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
#[derive(Debug, Clone)]
pub struct DocumentSettings {
    /// Whether document is protected
    protected: bool,
    /// Type of protection
    protection_type: Option<ProtectionType>,
    /// Whether to track revisions
    track_revisions: bool,
    /// Zoom percentage
    zoom_percent: Option<u32>,
}

/// Type of document protection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ProtectionType {
    /// No editing allowed
    ReadOnly,
    /// Only comments allowed
    Comments,
    /// Only tracked changes allowed
    TrackedChanges,
    /// Only form fields allowed
    Forms,
}

impl ProtectionType {
    /// Parse protection type from XML value.
    fn from_xml(s: &str) -> Option<Self> {
        match s {
            "readOnly" => Some(Self::ReadOnly),
            "comments" => Some(Self::Comments),
            "trackedChanges" => Some(Self::TrackedChanges),
            "forms" => Some(Self::Forms),
            _ => None,
        }
    }

    /// Get XML value for this protection type.
    pub const fn to_xml(self) -> &'static str {
        match self {
            Self::ReadOnly => "readOnly",
            Self::Comments => "comments",
            Self::TrackedChanges => "trackedChanges",
            Self::Forms => "forms",
        }
    }
}

impl DocumentSettings {
    /// Create a new DocumentSettings with default values.
    pub fn new() -> Self {
        Self {
            protected: false,
            protection_type: None,
            track_revisions: false,
            zoom_percent: None,
        }
    }

    /// Check if the document is protected.
    #[inline]
    pub fn is_protected(&self) -> bool {
        self.protected
    }

    /// Get the type of protection applied.
    #[inline]
    pub fn protection_type(&self) -> Option<ProtectionType> {
        self.protection_type
    }

    /// Check if track revisions is enabled.
    #[inline]
    pub fn track_revisions(&self) -> bool {
        self.track_revisions
    }

    /// Get the zoom percentage.
    #[inline]
    pub fn zoom_percent(&self) -> Option<u32> {
        self.zoom_percent
    }

    /// Extract settings from a settings.xml part.
    ///
    /// # Arguments
    ///
    /// * `part` - The settings part
    ///
    /// # Returns
    ///
    /// A DocumentSettings object
    pub(crate) fn extract_from_part(part: &dyn Part) -> Result<Self> {
        let xml_bytes = part.blob();
        let mut reader = Reader::from_reader(xml_bytes);
        reader.config_mut().trim_text(true);

        let mut settings = Self::new();

        loop {
            match reader.read_event() {
                Ok(Event::Empty(e)) | Ok(Event::Start(e)) => {
                    match e.local_name().as_ref() {
                        b"documentProtection" => {
                            settings.protected = true;

                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"edit" => {
                                        let val = String::from_utf8_lossy(&attr.value);
                                        settings.protection_type = ProtectionType::from_xml(&val);
                                    },
                                    b"enforcement" => {
                                        let val = String::from_utf8_lossy(&attr.value);
                                        // If enforcement is false, document is not actually protected
                                        if val == "false" || val == "0" {
                                            settings.protected = false;
                                        }
                                    },
                                    _ => {},
                                }
                            }
                        },
                        b"trackRevisions" => {
                            // Check for val attribute, or assume true if element exists
                            let mut has_val = false;
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"val" {
                                    has_val = true;
                                    let val = String::from_utf8_lossy(&attr.value);
                                    settings.track_revisions = val == "true" || val == "1";
                                }
                            }
                            if !has_val {
                                settings.track_revisions = true;
                            }
                        },
                        b"zoom" => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"percent" {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    settings.zoom_percent =
                                        atoi_simd::parse::<u32>(val.as_bytes()).ok();
                                }
                            }
                        },
                        _ => {},
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(settings)
    }
}

impl Default for DocumentSettings {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_settings_creation() {
        let settings = DocumentSettings::new();
        assert!(!settings.is_protected());
        assert!(settings.protection_type().is_none());
        assert!(!settings.track_revisions());
    }

    #[test]
    fn test_protection_type() {
        assert_eq!(
            ProtectionType::from_xml("readOnly"),
            Some(ProtectionType::ReadOnly)
        );
        assert_eq!(
            ProtectionType::from_xml("comments"),
            Some(ProtectionType::Comments)
        );
        assert_eq!(
            ProtectionType::from_xml("trackedChanges"),
            Some(ProtectionType::TrackedChanges)
        );
        assert_eq!(
            ProtectionType::from_xml("forms"),
            Some(ProtectionType::Forms)
        );
        assert_eq!(ProtectionType::from_xml("invalid"), None);
    }
}
