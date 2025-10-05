/// Constant values related to the Open Packaging Convention.
///
/// This module contains content type URIs (like MIME-types) that specify a part's format,
/// XML namespaces, and relationship types used in OPC packages.

/// Content type URIs (like MIME-types) that specify a part's format
pub mod content_type {
    // Image content types
    pub const BMP: &str = "image/bmp";
    pub const GIF: &str = "image/gif";
    pub const JPEG: &str = "image/jpeg";
    pub const PNG: &str = "image/png";
    pub const TIFF: &str = "image/tiff";
    pub const MS_PHOTO: &str = "image/vnd.ms-photo";
    pub const X_EMF: &str = "image/x-emf";
    pub const X_WMF: &str = "image/x-wmf";

    // DrawingML content types
    pub const DML_CHART: &str = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
    pub const DML_CHARTSHAPES: &str =
        "application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml";
    pub const DML_DIAGRAM_COLORS: &str =
        "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml";
    pub const DML_DIAGRAM_DATA: &str =
        "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml";
    pub const DML_DIAGRAM_LAYOUT: &str =
        "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml";
    pub const DML_DIAGRAM_STYLE: &str =
        "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml";

    // Office common content types
    pub const OFC_CUSTOM_PROPERTIES: &str =
        "application/vnd.openxmlformats-officedocument.custom-properties+xml";
    pub const OFC_CUSTOM_XML_PROPERTIES: &str =
        "application/vnd.openxmlformats-officedocument.customXmlProperties+xml";
    pub const OFC_DRAWING: &str = "application/vnd.openxmlformats-officedocument.drawing+xml";
    pub const OFC_EXTENDED_PROPERTIES: &str =
        "application/vnd.openxmlformats-officedocument.extended-properties+xml";
    pub const OFC_OLE_OBJECT: &str = "application/vnd.openxmlformats-officedocument.oleObject";
    pub const OFC_PACKAGE: &str = "application/vnd.openxmlformats-officedocument.package";
    pub const OFC_THEME: &str = "application/vnd.openxmlformats-officedocument.theme+xml";
    pub const OFC_THEME_OVERRIDE: &str =
        "application/vnd.openxmlformats-officedocument.themeOverride+xml";
    pub const OFC_VML_DRAWING: &str = "application/vnd.openxmlformats-officedocument.vmlDrawing";

    // OPC core content types
    pub const OPC_CORE_PROPERTIES: &str =
        "application/vnd.openxmlformats-package.core-properties+xml";
    pub const OPC_DIGITAL_SIGNATURE_CERTIFICATE: &str =
        "application/vnd.openxmlformats-package.digital-signature-certificate";
    pub const OPC_DIGITAL_SIGNATURE_ORIGIN: &str =
        "application/vnd.openxmlformats-package.digital-signature-origin";
    pub const OPC_DIGITAL_SIGNATURE_XMLSIGNATURE: &str =
        "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml";
    pub const OPC_RELATIONSHIPS: &str = "application/vnd.openxmlformats-package.relationships+xml";

    // WordprocessingML content types
    pub const WML_COMMENTS: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
    pub const WML_DOCUMENT: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    pub const WML_DOCUMENT_GLOSSARY: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml";
    pub const WML_DOCUMENT_MAIN: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
    pub const WML_ENDNOTES: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml";
    pub const WML_FONT_TABLE: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml";
    pub const WML_FOOTER: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
    pub const WML_FOOTNOTES: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml";
    pub const WML_HEADER: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
    pub const WML_NUMBERING: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
    pub const WML_SETTINGS: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";
    pub const WML_STYLES: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
    pub const WML_WEB_SETTINGS: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml";

    // SpreadsheetML content types
    pub const SML_SHEET: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    pub const SML_SHEET_MAIN: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
    pub const SML_WORKSHEET: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
    pub const SML_STYLES: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
    pub const SML_SHARED_STRINGS: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";

    // PresentationML content types
    pub const PML_PRESENTATION_MAIN: &str =
        "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
    pub const PML_SLIDE: &str =
        "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
    pub const PML_SLIDE_LAYOUT: &str =
        "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml";
    pub const PML_SLIDE_MASTER: &str =
        "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml";

    // Generic XML
    pub const XML: &str = "application/xml";
}

/// XML namespace URIs used in OPC packages
pub mod namespace {
    /// DrawingML wordprocessing drawing namespace
    pub const DML_WORDPROCESSING_DRAWING: &str =
        "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

    /// Office relationships namespace
    pub const OFC_RELATIONSHIPS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// OPC relationships namespace
    pub const OPC_RELATIONSHIPS: &str =
        "http://schemas.openxmlformats.org/package/2006/relationships";

    /// OPC content types namespace
    pub const OPC_CONTENT_TYPES: &str =
        "http://schemas.openxmlformats.org/package/2006/content-types";

    /// WordprocessingML main namespace
    pub const WML_MAIN: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
}

/// Open XML relationship target modes
pub mod target_mode {
    /// Internal relationship target mode (default)
    pub const INTERNAL: &str = "Internal";

    /// External relationship target mode (e.g., hyperlinks to external URLs)
    pub const EXTERNAL: &str = "External";
}

/// Relationship type URIs used in OPC packages
pub mod relationship_type {
    // Core relationships
    pub const CORE_PROPERTIES: &str =
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
    pub const EXTENDED_PROPERTIES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
    pub const CUSTOM_PROPERTIES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";
    pub const THUMBNAIL: &str =
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail";

    // Office document
    pub const OFFICE_DOCUMENT: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

    // Document parts
    pub const COMMENTS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
    pub const ENDNOTES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";
    pub const FONT_TABLE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable";
    pub const FOOTER: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
    pub const FOOTNOTES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
    pub const HEADER: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
    pub const NUMBERING: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
    pub const SETTINGS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
    pub const STYLES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
    pub const WEB_SETTINGS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings";

    // Images and media
    pub const IMAGE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
    pub const AUDIO: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio";
    pub const VIDEO: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video";

    // Chart and drawing
    pub const CHART: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
    pub const DRAWING: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
    pub const VML_DRAWING: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";

    // Theme
    pub const THEME: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
    pub const THEME_OVERRIDE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/themeOverride";

    // External links
    pub const HYPERLINK: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    pub const OLE_OBJECT: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject";
    pub const PACKAGE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package";
}
