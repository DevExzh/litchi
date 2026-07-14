//! Bundle metadata values and application detection.

use std::collections::HashMap;
use std::path::Path;

use crate::Result;

/// Metadata associated with an iWork bundle
#[derive(Debug, Clone, Default)]
pub struct BundleMetadata {
    /// Whether Properties.plist exists
    pub has_properties: bool,
    /// Whether BuildVersionHistory.plist exists
    pub has_build_version_history: bool,
    /// Whether DocumentIdentifier exists
    pub has_document_identifier: bool,
    /// Application type detected from the bundle
    pub detected_application: Option<String>,
    /// Parsed properties from Properties.plist
    pub properties: HashMap<String, PropertyValue>,
    /// Build version history
    pub build_versions: Vec<String>,
    /// Document identifier
    pub document_id: Option<String>,
}

/// Represents a property value from plist
#[derive(Debug, Clone)]
pub enum PropertyValue {
    /// String value
    String(String),
    /// Integer value
    Integer(i64),
    /// Real/float value
    Real(f64),
    /// Boolean value
    Boolean(bool),
    /// Date value
    Date(String),
    /// Array of values
    Array(Vec<PropertyValue>),
    /// Dictionary of values
    Dictionary(HashMap<String, PropertyValue>),
}

impl BundleMetadata {
    /// Get a summary of the metadata
    pub fn summary(&self) -> String {
        format!(
            "Properties: {}, BuildVersion: {}, DocumentID: {}, App: {}",
            self.has_properties,
            self.has_build_version_history,
            self.has_document_identifier,
            self.detected_application.as_deref().unwrap_or("unknown")
        )
    }

    /// Get a property value as a string
    pub fn get_property_string(&self, key: &str) -> Option<String> {
        match self.properties.get(key)? {
            PropertyValue::String(s) => Some(s.clone()),
            PropertyValue::Integer(i) => Some(i.to_string()),
            PropertyValue::Real(r) => Some(r.to_string()),
            PropertyValue::Boolean(b) => Some(b.to_string()),
            PropertyValue::Date(d) => Some(d.clone()),
            _ => None,
        }
    }

    /// Get a property value as an integer
    pub fn get_property_int(&self, key: &str) -> Option<i64> {
        match self.properties.get(key)? {
            PropertyValue::Integer(i) => Some(*i),
            _ => None,
        }
    }

    /// Get a property value as a boolean
    pub fn get_property_bool(&self, key: &str) -> Option<bool> {
        match self.properties.get(key)? {
            PropertyValue::Boolean(b) => Some(*b),
            _ => None,
        }
    }

    /// Get the document identifier
    pub fn document_identifier(&self) -> Option<&str> {
        self.document_id.as_deref()
    }

    /// Get the build versions
    pub fn build_version_history(&self) -> &[String] {
        &self.build_versions
    }

    /// Get the latest build version
    pub fn latest_build_version(&self) -> Option<&str> {
        self.build_versions.last().map(|s| s.as_str())
    }
}

/// Detect the application type from a bundle path
pub fn detect_application_type<P: AsRef<Path>>(bundle_path: P) -> Result<String> {
    let path = bundle_path.as_ref();

    // Check file extension or directory structure
    if let Some(extension) = path.extension() {
        match extension.to_str() {
            Some("pages") => return Ok("Pages".to_string()),
            Some("key") => return Ok("Keynote".to_string()),
            Some("numbers") => return Ok("Numbers".to_string()),
            _ => {},
        }
    }

    // Check for application-specific files in Index.zip
    if path.is_dir() {
        let index_zip = path.join("Index.zip");
        if index_zip.exists() {
            // This would require opening the zip and checking for app-specific files
            // For now, return "Unknown"
        }
    }

    Ok("Unknown".to_string())
}
