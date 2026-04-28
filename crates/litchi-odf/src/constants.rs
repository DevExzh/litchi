//! ODF constants, MIME types, and XML tags.
//!
//! This module provides comprehensive constants for OpenDocument Format (ODF) parsing and creation.
//! Based on reference implementations from odfdo and odfpy libraries.
//!
//! # Implementation Status
//!
//! ✅ COMPLETED: MIME type constants
//! ✅ COMPLETED: File extension mapping
//! ✅ COMPLETED: Standard ODF part paths
//! ✅ COMPLETED: Office version constant
//! ✅ COMPLETED: Presentation class constants
//!
//! # References
//!
//! - odfdo: `3rdparty/odfdo/src/odfdo/const.py`
//! - odfpy: `3rdparty/odfpy/odf/namespaces.py`

use phf::{Map, phf_map};

/// ODF specification version
pub const OFFICE_VERSION: &str = "1.2";

// ============================================================================
// MIME TYPES
// ============================================================================
// Reference: odfdo/const.py lines 52-67

/// MIME type for OpenDocument Text (.odt)
pub const ODF_TEXT: &str = "application/vnd.oasis.opendocument.text";

/// MIME type for OpenDocument Text Template (.ott)
pub const ODF_TEXT_TEMPLATE: &str = "application/vnd.oasis.opendocument.text-template";

/// MIME type for OpenDocument Spreadsheet (.ods)
pub const ODF_SPREADSHEET: &str = "application/vnd.oasis.opendocument.spreadsheet";

/// MIME type for OpenDocument Spreadsheet Template (.ots)
pub const ODF_SPREADSHEET_TEMPLATE: &str =
    "application/vnd.oasis.opendocument.spreadsheet-template";

/// MIME type for OpenDocument Presentation (.odp)
pub const ODF_PRESENTATION: &str = "application/vnd.oasis.opendocument.presentation";

/// MIME type for OpenDocument Presentation Template (.otp)
pub const ODF_PRESENTATION_TEMPLATE: &str =
    "application/vnd.oasis.opendocument.presentation-template";

/// MIME type for OpenDocument Drawing (.odg)
pub const ODF_DRAWING: &str = "application/vnd.oasis.opendocument.graphics";

/// MIME type for OpenDocument Drawing Template (.otg)
pub const ODF_DRAWING_TEMPLATE: &str = "application/vnd.oasis.opendocument.graphics-template";

/// MIME type for OpenDocument Chart (.odc)
pub const ODF_CHART: &str = "application/vnd.oasis.opendocument.chart";

/// MIME type for OpenDocument Chart Template (.otc)
pub const ODF_CHART_TEMPLATE: &str = "application/vnd.oasis.opendocument.chart-template";

/// MIME type for OpenDocument Image (.odi)
pub const ODF_IMAGE: &str = "application/vnd.oasis.opendocument.image";

/// MIME type for OpenDocument Image Template (.oti)
pub const ODF_IMAGE_TEMPLATE: &str = "application/vnd.oasis.opendocument.image-template";

/// MIME type for OpenDocument Formula (.odf)
pub const ODF_FORMULA: &str = "application/vnd.oasis.opendocument.formula";

/// MIME type for OpenDocument Formula Template (.otf)
pub const ODF_FORMULA_TEMPLATE: &str = "application/vnd.oasis.opendocument.formula-template";

/// MIME type for OpenDocument Master (.odm)
pub const ODF_MASTER: &str = "application/vnd.oasis.opendocument.text-master";

/// MIME type for OpenDocument Web (.oth)
pub const ODF_WEB: &str = "application/vnd.oasis.opendocument.text-web";

// ============================================================================
// FILE EXTENSIONS TO MIME TYPE MAPPING
// ============================================================================
// Reference: odfdo/const.py lines 70-87
// Using phf for compile-time perfect hash map - zero runtime overhead

/// File extension to MIME type mapping (compile-time perfect hash map)
pub static ODF_EXTENSIONS: Map<&'static str, &'static str> = phf_map! {
    "odt" => ODF_TEXT,
    "ott" => ODF_TEXT_TEMPLATE,
    "ods" => ODF_SPREADSHEET,
    "ots" => ODF_SPREADSHEET_TEMPLATE,
    "odp" => ODF_PRESENTATION,
    "otp" => ODF_PRESENTATION_TEMPLATE,
    "odg" => ODF_DRAWING,
    "otg" => ODF_DRAWING_TEMPLATE,
    "odc" => ODF_CHART,
    "otc" => ODF_CHART_TEMPLATE,
    "odi" => ODF_IMAGE,
    "oti" => ODF_IMAGE_TEMPLATE,
    "odf" => ODF_FORMULA,
    "otf" => ODF_FORMULA_TEMPLATE,
    "odm" => ODF_MASTER,
    "oth" => ODF_WEB,
};

/// MIME type to file extension mapping (compile-time perfect hash map)
pub static ODF_MIMETYPES: Map<&'static str, &'static str> = phf_map! {
    "application/vnd.oasis.opendocument.text" => "odt",
    "application/vnd.oasis.opendocument.text-template" => "ott",
    "application/vnd.oasis.opendocument.spreadsheet" => "ods",
    "application/vnd.oasis.opendocument.spreadsheet-template" => "ots",
    "application/vnd.oasis.opendocument.presentation" => "odp",
    "application/vnd.oasis.opendocument.presentation-template" => "otp",
    "application/vnd.oasis.opendocument.graphics" => "odg",
    "application/vnd.oasis.opendocument.graphics-template" => "otg",
    "application/vnd.oasis.opendocument.chart" => "odc",
    "application/vnd.oasis.opendocument.chart-template" => "otc",
    "application/vnd.oasis.opendocument.image" => "odi",
    "application/vnd.oasis.opendocument.image-template" => "oti",
    "application/vnd.oasis.opendocument.formula" => "odf",
    "application/vnd.oasis.opendocument.formula-template" => "otf",
    "application/vnd.oasis.opendocument.text-master" => "odm",
    "application/vnd.oasis.opendocument.text-web" => "oth",
};

// ============================================================================
// STANDARD ODF PARTS PATHS
// ============================================================================
// Reference: odfdo/const.py lines 110-118

/// Path to content.xml (main document content)
pub const ODF_CONTENT: &str = "content.xml";

/// Path to meta.xml (document metadata)
pub const ODF_META: &str = "meta.xml";

/// Path to settings.xml (application settings)
pub const ODF_SETTINGS: &str = "settings.xml";

/// Path to styles.xml (document styles)
pub const ODF_STYLES: &str = "styles.xml";

/// Path to manifest.xml (package manifest)
pub const ODF_MANIFEST: &str = "META-INF/manifest.xml";

/// Name of the manifest file
pub const ODF_MANIFEST_NAME: &str = "manifest.xml";

/// Path to manifest.rdf (RDF metadata)
pub const ODF_MANIFEST_RDF: &str = "manifest.rdf";

/// MIME type for RDF manifest
pub const ODF_MANIFEST_RDF_TYPE: &str = "application/rdf+xml";

/// Standard parts in the ODF container
pub const ODF_PARTS: [&str; 5] = ["content", "meta", "settings", "styles", "manifest"];

// ============================================================================
// PRESENTATION CLASSES (for layout)
// ============================================================================
// Reference: odfdo/const.py lines 175-188

/// Presentation class constants for slide layouts
pub const ODF_CLASSES: [&str; 12] = [
    "title", "outline", "subtitle", "text", "graphic", "object", "chart", "table", "orgchart",
    "page", "notes", "handout",
];

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/// Get MIME type from file extension
///
/// # Arguments
///
/// * `extension` - File extension (without dot, e.g., "odt")
///
/// # Returns
///
/// The corresponding MIME type if recognized, None otherwise.
///
/// # Examples
///
/// ```
/// use litchi::odf::constants::get_mime_type_from_extension;
///
/// let mime = get_mime_type_from_extension("odt");
/// assert_eq!(mime, Some("application/vnd.oasis.opendocument.text"));
/// ```
#[inline]
pub fn get_mime_type_from_extension(extension: &str) -> Option<&'static str> {
    ODF_EXTENSIONS.get(extension).copied()
}

/// Get file extension from MIME type
///
/// # Arguments
///
/// * `mime_type` - MIME type string
///
/// # Returns
///
/// The corresponding file extension if recognized, None otherwise.
///
/// # Examples
///
/// ```
/// use litchi::odf::constants::get_extension_from_mime_type;
///
/// let ext = get_extension_from_mime_type("application/vnd.oasis.opendocument.text");
/// assert_eq!(ext, Some("odt"));
/// ```
#[inline]
pub fn get_extension_from_mime_type(mime_type: &str) -> Option<&'static str> {
    ODF_MIMETYPES.get(mime_type).copied()
}

/// Check if a given extension is a valid ODF extension
///
/// # Arguments
///
/// * `extension` - File extension (without dot)
///
/// # Returns
///
/// `true` if the extension is a valid ODF extension, `false` otherwise.
///
/// # Examples
///
/// ```
/// use litchi::odf::constants::is_odf_extension;
///
/// assert!(is_odf_extension("odt"));
/// assert!(is_odf_extension("ods"));
/// assert!(!is_odf_extension("txt"));
/// ```
#[inline]
pub fn is_odf_extension(extension: &str) -> bool {
    ODF_EXTENSIONS.contains_key(extension)
}

/// Check if a given MIME type is a valid ODF MIME type
///
/// # Arguments
///
/// * `mime_type` - MIME type string
///
/// # Returns
///
/// `true` if the MIME type is a valid ODF MIME type, `false` otherwise.
///
/// # Examples
///
/// ```
/// use litchi::odf::constants::is_odf_mime_type;
///
/// assert!(is_odf_mime_type("application/vnd.oasis.opendocument.text"));
/// assert!(!is_odf_mime_type("text/plain"));
/// ```
#[inline]
pub fn is_odf_mime_type(mime_type: &str) -> bool {
    ODF_MIMETYPES.contains_key(mime_type)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_extension_to_mime() {
        assert_eq!(get_mime_type_from_extension("odt"), Some(ODF_TEXT));
        assert_eq!(get_mime_type_from_extension("ods"), Some(ODF_SPREADSHEET));
        assert_eq!(get_mime_type_from_extension("odp"), Some(ODF_PRESENTATION));
        assert_eq!(get_mime_type_from_extension("unknown"), None);
    }

    #[test]
    fn test_mime_to_extension() {
        assert_eq!(get_extension_from_mime_type(ODF_TEXT), Some("odt"));
        assert_eq!(get_extension_from_mime_type(ODF_SPREADSHEET), Some("ods"));
        assert_eq!(get_extension_from_mime_type(ODF_PRESENTATION), Some("odp"));
        assert_eq!(get_extension_from_mime_type("unknown"), None);
    }

    #[test]
    fn test_is_odf_extension() {
        assert!(is_odf_extension("odt"));
        assert!(is_odf_extension("ods"));
        assert!(is_odf_extension("odp"));
        assert!(!is_odf_extension("txt"));
        assert!(!is_odf_extension("docx"));
    }

    #[test]
    fn test_is_odf_mime_type() {
        assert!(is_odf_mime_type(ODF_TEXT));
        assert!(is_odf_mime_type(ODF_SPREADSHEET));
        assert!(is_odf_mime_type(ODF_PRESENTATION));
        assert!(!is_odf_mime_type("text/plain"));
        assert!(!is_odf_mime_type("application/pdf"));
    }

    #[test]
    fn test_standard_paths() {
        assert_eq!(ODF_CONTENT, "content.xml");
        assert_eq!(ODF_META, "meta.xml");
        assert_eq!(ODF_SETTINGS, "settings.xml");
        assert_eq!(ODF_STYLES, "styles.xml");
        assert_eq!(ODF_MANIFEST, "META-INF/manifest.xml");
    }

    #[test]
    fn test_office_version() {
        assert_eq!(OFFICE_VERSION, "1.2");
    }

    #[test]
    fn test_presentation_classes() {
        assert!(ODF_CLASSES.contains(&"title"));
        assert!(ODF_CLASSES.contains(&"subtitle"));
        assert!(ODF_CLASSES.contains(&"chart"));
        assert_eq!(ODF_CLASSES.len(), 12);
    }
}
