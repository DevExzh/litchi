/// Document template module.
///
/// Provides minimal valid templates for creating new Word documents.
/// These templates contain the bare minimum structure required for a valid .docx file.
/// Generate a minimal valid document.xml content.
///
/// Creates an empty document with a single section definition.
pub fn default_document_xml() -> &'static str {
    include_str!("resources/document.xml")
}

/// Generate default styles.xml content
///
/// Uses theme references for fonts instead of direct font names to ensure compatibility.
pub fn default_styles_xml() -> &'static str {
    include_str!("resources/styles.xml")
}

/// Generate default settings.xml content
pub fn default_settings_xml() -> &'static str {
    include_str!("resources/settings.xml")
}

/// Generate a minimal valid fontTable.xml content.
pub fn default_font_table_xml() -> &'static str {
    include_str!("resources/fontTable.xml")
}

/// Generate a minimal valid webSettings.xml content.
pub fn default_web_settings_xml() -> &'static str {
    include_str!("resources/webSettings.xml")
}

/// Generate a minimal valid core.xml (core properties) content.
pub fn default_core_props_xml() -> &'static str {
    include_str!("resources/docProps/core.xml")
}

/// Generate a minimal valid app.xml (extended properties) content.
pub fn default_app_props_xml() -> &'static str {
    include_str!("resources/docProps/app.xml")
}

/// Generate a minimal valid theme1.xml content.
///
/// Defines the Office theme with color scheme and font scheme.
pub fn default_theme_xml() -> &'static str {
    include_str!("resources/theme/theme1.xml")
}
