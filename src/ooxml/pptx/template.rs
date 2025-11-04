//! Presentation template module.
//!
//! Provides minimal valid templates for creating new PowerPoint presentations.
//! These templates contain the bare minimum structure required for a valid .pptx file.
//! Generate valid presentation.xml content (based on python-pptx template).

use xml_minifier::minified_xml;

/// Creates an empty presentation with no slides but complete text styling.
pub fn default_presentation_xml() -> &'static str {
    minified_xml!("resources/presentation.xml")
}

/// Generate a comprehensive valid slideMaster.xml content (based on python-pptx template).
///
/// This includes:
/// - Proper placeholder shapes (title, body, date, footer, slide number)
/// - Complete text styles with 9 levels for title, body, and other styles
/// - Color mapping
/// - Slide layout ID list
pub fn default_slide_master_xml() -> &'static str {
    minified_xml!("resources/slideMasters/slideMaster1.xml")
}

/// Generate slide layout 1 XML (Title Slide)
pub fn slide_layout_1_xml() -> &'static str {
    minified_xml!("resources/slideLayouts/slideLayout1.xml")
}

/// Generate slide layout 2 XML (Title and Content)
pub fn slide_layout_2_xml() -> &'static str {
    minified_xml!("resources/slideLayouts/slideLayout2.xml")
}

/// Generate slide layout 3 XML (Section Header)
pub fn slide_layout_3_xml() -> &'static str {
    minified_xml!("resources/slideLayouts/slideLayout3.xml")
}

/// Generate slide layout 4 XML (Two Content)
pub fn slide_layout_4_xml() -> &'static str {
    minified_xml!("resources/slideLayouts/slideLayout4.xml")
}

/// Generate slide layout 5 XML (Comparison)
pub fn slide_layout_5_xml() -> &'static str {
    minified_xml!("resources/slideLayouts/slideLayout5.xml")
}

/// Generate slide layout 6 XML (Title Only)
pub fn slide_layout_6_xml() -> &'static str {
    minified_xml!("resources/slideLayouts/slideLayout6.xml")
}

/// Generate slide layout 7 XML (Blank)
pub fn slide_layout_7_xml() -> &'static str {
    minified_xml!("resources/slideLayouts/slideLayout7.xml")
}

/// Generate slide layout 8 XML (Content with Caption)
pub fn slide_layout_8_xml() -> &'static str {
    minified_xml!("resources/slideLayouts/slideLayout8.xml")
}

/// Generate slide layout 9 XML (Picture with Caption)
pub fn slide_layout_9_xml() -> &'static str {
    minified_xml!("resources/slideLayouts/slideLayout9.xml")
}

/// Generate slide layout 10 XML (Title and Vertical Text)
pub fn slide_layout_10_xml() -> &'static str {
    minified_xml!("resources/slideLayouts/slideLayout10.xml")
}

/// Generate slide layout 11 XML (Vertical Title and Text)
pub fn slide_layout_11_xml() -> &'static str {
    minified_xml!("resources/slideLayouts/slideLayout11.xml")
}

/// Get all slide layout XMLs as a vector
pub fn all_slide_layouts() -> Vec<&'static str> {
    vec![
        slide_layout_1_xml(),
        slide_layout_2_xml(),
        slide_layout_3_xml(),
        slide_layout_4_xml(),
        slide_layout_5_xml(),
        slide_layout_6_xml(),
        slide_layout_7_xml(),
        slide_layout_8_xml(),
        slide_layout_9_xml(),
        slide_layout_10_xml(),
        slide_layout_11_xml(),
    ]
}

/// Generate notes master XML
pub fn default_notes_master_xml() -> &'static str {
    minified_xml!("resources/notesMaster.xml")
}

/// Generate a minimal valid theme.xml content.
pub fn default_theme_xml() -> &'static str {
    minified_xml!("resources/theme/theme1.xml")
}

/// Generate a minimal valid tableStyles.xml content.
pub fn default_table_styles_xml() -> &'static str {
    minified_xml!("resources/tableStyles.xml")
}

/// Generate a minimal valid viewProps.xml content.
pub fn default_view_props_xml() -> &'static str {
    minified_xml!("resources/viewProps.xml")
}

/// Generate a minimal valid presProps.xml content.
pub fn default_pres_props_xml() -> &'static str {
    minified_xml!("resources/presProps.xml")
}

/// Generate a minimal valid core.xml (core properties) content.
pub fn default_core_props_xml() -> &'static str {
    minified_xml!("resources/docProps/core.xml")
}

/// Generate a minimal valid app.xml (extended properties) content.
pub fn default_app_props_xml() -> &'static str {
    minified_xml!("resources/docProps/app.xml")
}
