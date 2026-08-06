//! Typed SpreadsheetML chartsheet semantic model.
//!
//! The root model is kept deliberately small. Its schema children live in
//! focused modules so that semantic values do not get mixed with the OPC
//! resources owned by [`super::package`]. The parent module re-exports the
//! children to keep the `chart_sheet::*` facade flat.

mod extensions;
mod metadata;
mod publishing;
mod views;

pub use extensions::*;
pub use metadata::*;
pub use publishing::*;
pub use views::*;

pub(super) const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
pub(super) const STRICT_SML: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
pub(super) const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
pub(super) const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
pub(super) const XDR: &str = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
pub(super) const STRICT_XDR: &str = "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing";
pub(super) const CHART: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
pub(super) const STRICT_CHART: &str = "http://purl.oclc.org/ooxml/drawingml/chart";
pub(super) const DRAWING_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
pub(super) const STRICT_DRAWING_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/drawing";
pub(super) const CHART_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
pub(super) const STRICT_CHART_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/chart";
pub(super) const CHARTSHEET_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
pub(super) const STRICT_CHARTSHEET_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/chartsheet";
pub(super) const IMAGE_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
pub(super) const STRICT_IMAGE_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/image";
pub(super) const CHART_USER_SHAPES_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartUserShapes";
pub(super) const STRICT_CHART_USER_SHAPES_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/chartUserShapes";
pub(super) const THEME_OVERRIDE_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/themeOverride";
pub(super) const STRICT_THEME_OVERRIDE_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/themeOverride";
pub(super) const PACKAGE_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package";
pub(super) const STRICT_PACKAGE_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/package";
pub(super) const VML_DRAWING_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
pub(super) const STRICT_VML_DRAWING_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/vmlDrawing";
pub(super) const PRINTER_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings";
pub(super) const STRICT_PRINTER_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/printerSettings";

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Conformance {
    Transitional,
    Strict,
}

impl Conformance {
    pub fn sml(self) -> &'static str {
        match self {
            Self::Transitional => SML,
            Self::Strict => STRICT_SML,
        }
    }

    pub fn rel(self) -> &'static str {
        match self {
            Self::Transitional => REL,
            Self::Strict => STRICT_REL,
        }
    }

    pub fn xdr(self) -> &'static str {
        match self {
            Self::Transitional => XDR,
            Self::Strict => STRICT_XDR,
        }
    }

    pub fn chart(self) -> &'static str {
        match self {
            Self::Transitional => CHART,
            Self::Strict => STRICT_CHART,
        }
    }

    pub fn chartsheet_rel(self) -> &'static str {
        match self {
            Self::Transitional => CHARTSHEET_REL,
            Self::Strict => STRICT_CHARTSHEET_REL,
        }
    }

    pub fn drawing_rel(self) -> &'static str {
        match self {
            Self::Transitional => DRAWING_REL,
            Self::Strict => STRICT_DRAWING_REL,
        }
    }

    pub fn chart_rel(self) -> &'static str {
        match self {
            Self::Transitional => CHART_REL,
            Self::Strict => STRICT_CHART_REL,
        }
    }

    pub fn image_rel(self) -> &'static str {
        match self {
            Self::Transitional => IMAGE_REL,
            Self::Strict => STRICT_IMAGE_REL,
        }
    }

    pub fn chart_user_shapes_rel(self) -> &'static str {
        match self {
            Self::Transitional => CHART_USER_SHAPES_REL,
            Self::Strict => STRICT_CHART_USER_SHAPES_REL,
        }
    }

    pub fn theme_override_rel(self) -> &'static str {
        match self {
            Self::Transitional => THEME_OVERRIDE_REL,
            Self::Strict => STRICT_THEME_OVERRIDE_REL,
        }
    }

    pub fn package_rel(self) -> &'static str {
        match self {
            Self::Transitional => PACKAGE_REL,
            Self::Strict => STRICT_PACKAGE_REL,
        }
    }

    pub fn vml_drawing_rel(self) -> &'static str {
        match self {
            Self::Transitional => VML_DRAWING_REL,
            Self::Strict => STRICT_VML_DRAWING_REL,
        }
    }

    pub fn printer_rel(self) -> &'static str {
        match self {
            Self::Transitional => PRINTER_REL,
            Self::Strict => STRICT_PRINTER_REL,
        }
    }
}

/// The typed semantic contents of one SpreadsheetML `chartsheet` part.
#[derive(Debug, Clone, PartialEq)]
pub struct Chart {
    pub properties: Option<Properties>,
    pub views: Vec<View>,
    pub protection: Option<Protection>,
    /// `None` preserves absence; a present collection must be non-empty.
    pub custom_views: Option<Vec<CustomView>>,
    pub margins: Option<Margins>,
    pub page_setup: Option<PageSetup>,
    pub header_footer: Option<HeaderFooter>,
    pub drawing_relationship_id: String,
    pub legacy_drawing_relationship_id: Option<String>,
    pub legacy_header_footer_drawing_relationship_id: Option<String>,
    /// Relationship for the optional tiled chartsheet background image.
    pub background_picture_relationship_id: Option<String>,
    /// Inert web-publishing metadata. Destinations and sources are never resolved or accessed.
    pub web_publish_items: Option<WebPublishItems>,
    /// Inert, canonicalized wildcard extension markup. Payload semantics are never interpreted.
    pub extension_list: Option<ExtensionList>,
}
